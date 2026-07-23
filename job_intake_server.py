"""Loopback HTTPS listener that lets an Outlook add-in file a job intake.

Bound to 127.0.0.1 only and guarded by a bearer token, this exists so the
Phase 3 task pane can POST a job number plus the email's base64 DXF
attachments. It owns no intake logic of its own: decoding attachments to a
temp dir and calling ``job_intake_service.create_intake(..., source="outlook")``
is the entire request body. Any behavior change belongs in the service, where
the desktop page picks it up for free.

**No RADAN work happens here.** Cloning the RPD, importing parts, and sending
blocks stay explicit desktop actions - RADAN COM automation is slow and
interactive, and blocking a listener thread on it would hang the task pane.
The request creates the folder and the registry entry, then returns; the
intake shows up in the desktop queue on its next poll.
"""

from __future__ import annotations

import base64
import binascii
import hmac
import logging
import os
from pathlib import Path
import secrets
import shutil
import tempfile
import threading
from typing import Any

import job_intake_registry

from flask import Flask, Response, jsonify, render_template, request, send_from_directory

import job_intake_service
from job_intake_service import JobIntakeError
from job_intake_tls import ensure_loopback_certificate
from paths import APP_DIR


_logger = logging.getLogger(__name__)

HOST = "127.0.0.1"
DEFAULT_PORT = 8790
TOKEN_PATH = APP_DIR / "_runtime" / "job_intake_api_token.key"

# Outlook attachments arrive base64-encoded in JSON, so the decoded payload is
# ~1.33x smaller than the request. A DXF set for one job is far under this;
# the cap just stops a malformed request from exhausting memory.
MAX_CONTENT_LENGTH = 64 * 1024 * 1024

# Only these are accepted off the wire. DXFs are the actual work; PDFs are
# kept because the PO and drawing-print scrapes run against them; .csv/.xlsx
# because a BOM can arrive attached rather than at a W: path, and those are
# exactly the two extensions inventor_to_radan accepts
# (config.SUPPORTED_BOM_EXTENSIONS). Anything else in the email (images,
# signatures, .msg) is ignored rather than copied onto L:.
ALLOWED_SUFFIXES = (".dxf", ".pdf", ".csv", ".xlsx")

ADDIN_STATIC_DIR = APP_DIR / "static" / "job_intake_addin"

# Office resolves the manifest's URLs literally, so the hostname the add-in is
# registered with has to be one the certificate covers. Both loopback
# spellings are in the SAN; "localhost" is the default because that is what
# Office's own tooling and docs use throughout.
DEFAULT_ADDIN_HOST = "localhost"
ALLOWED_ADDIN_HOSTS = ("localhost", "127.0.0.1")


def listener_port() -> int:
    raw = os.environ.get("ODD_JOB_INTAKE_PORT", "").strip()
    try:
        return int(raw) if raw else DEFAULT_PORT
    except ValueError:
        _logger.warning("ODD_JOB_INTAKE_PORT=%r is not a number; using %d.", raw, DEFAULT_PORT)
        return DEFAULT_PORT


def ensure_api_token() -> str:
    """Read the shared secret, creating it on first run.

    Written 0600-ish by virtue of living under _runtime (which is gitignored);
    the task pane never asks the user to type it - Phase 3 injects it into the
    rendered page server-side.
    """
    try:
        existing = TOKEN_PATH.read_text(encoding="utf-8").strip()
        if existing:
            return existing
    except (FileNotFoundError, OSError):
        pass
    token = secrets.token_urlsafe(32)
    TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_PATH.write_text(token, encoding="utf-8")
    return token


def _safe_attachment_name(raw_name: str) -> str:
    """Reduce a client-supplied name to a bare filename with an allowed suffix.

    The name comes from an email attachment, so it is untrusted: strip any
    directory components (defeating ``..\\..\\`` traversal) before it is ever
    joined to a path under L:.
    """
    name = Path(str(raw_name or "").strip().replace("\\", "/")).name
    if not name or name in {".", ".."}:
        raise JobIntakeError(f"Attachment has an unusable filename: {raw_name!r}")
    if not name.casefold().endswith(ALLOWED_SUFFIXES):
        raise JobIntakeError(
            f"{name}: only {', '.join(ALLOWED_SUFFIXES)} attachments are accepted."
        )
    return name


def _decode_attachments(payload: Any, target_dir: Path) -> list[Path]:
    if not isinstance(payload, list) or not payload:
        raise JobIntakeError("Send at least one attachment as {name, contentBytes}.")
    written: list[Path] = []
    seen: set[str] = set()
    for item in payload:
        if not isinstance(item, dict):
            raise JobIntakeError("Each attachment must be an object with name and contentBytes.")
        name = _safe_attachment_name(item.get("name", ""))
        if name.casefold() in seen:
            raise JobIntakeError(f"{name}: the same attachment name was sent twice.")
        seen.add(name.casefold())
        try:
            content = base64.b64decode(str(item.get("contentBytes", "")), validate=True)
        except (binascii.Error, ValueError) as exc:
            raise JobIntakeError(f"{name}: contentBytes is not valid base64.") from exc
        if not content:
            raise JobIntakeError(f"{name}: the attachment is empty.")
        destination = target_dir / name
        destination.write_bytes(content)
        written.append(destination)
    return written


def _finish_intake(
    entry: dict[str, Any], files: list[Path], email_body: str, staging: str
) -> None:
    """The slow half of an intake, off the request thread.

    Copies the work files in, scrapes them and registers the finished entry.
    Failures are recorded on the entry rather than raised: nobody is waiting
    on this call, and a job whose folder exists but whose parts didn't resolve
    is still better surfaced in the queue than lost.

    Both endings must reach a terminal state. The macro polls until it sees
    one, so a failure left as `running` costs the user the whole five-minute
    timeout and then reports "still working" about a job that is already dead.
    """
    key = str(entry.get("key", ""))
    try:
        job_intake_registry.set_state(key, job_intake_registry.STATE_RUNNING)
    except ValueError:
        _logger.error("Could not mark %s as running.", key)
    try:
        # complete_intake writes STATE_SUCCEEDED itself, last, once everything
        # else about the entry has been saved.
        job_intake_service.complete_intake(entry, files, email_body=email_body)
        _logger.info("Completed intake %s (%d parts).", key, len(entry.get("material_qty", [])))
    except Exception as exc:
        _logger.exception("Intake %s failed after being queued.", key)
        # The entry already exists - begin_intake claimed it - so the failure
        # is recorded on it and shows up in the desktop queue.
        try:
            job_intake_registry.set_state(
                key,
                job_intake_registry.STATE_FAILED,
                status=job_intake_registry.STATUS_ERROR,
                error=f"Intake failed after queuing: {exc}",
            )
        except ValueError:
            _logger.error("Could not record the failure for %s.", key)
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def create_app(
    token: str | None = None,
    port: int | None = None,
    *,
    run_in_background: bool = True,
) -> Flask:
    """The listener.

    `run_in_background` is what makes the request return before the slow half
    finishes. Tests turn it off so a POST is complete when it returns; leaving
    it on there would mean racing a thread for every assertion.
    """
    # Pin root/template/static paths explicitly instead of letting Flask infer
    # them from __name__. When master_app embeds this repo it loads the modules
    # and then removes them from sys.modules, so Flask's inference falls back
    # to the host's working directory and every render_template 500s - while
    # importing this module normally (as the tests do) works fine. Deriving
    # from APP_DIR is correct in both cases.
    app = Flask(
        __name__,
        root_path=str(APP_DIR),
        template_folder=str(APP_DIR / "templates"),
        static_folder=str(APP_DIR / "static"),
    )
    app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

    # Any intake still marked queued/running belongs to a worker thread that
    # died with the previous process - nothing will ever finish it. Closing
    # them out here is what keeps a poller from waiting on a job that stopped
    # when the app was last closed.
    try:
        stranded = job_intake_registry.fail_interrupted_entries()
        if stranded:
            _logger.warning(
                "Closed out %d intake(s) interrupted by a restart: %s",
                len(stranded),
                ", ".join(stranded),
            )
    except Exception:
        _logger.exception("Could not check for interrupted intakes.")

    expected_token = token if token is not None else ensure_api_token()
    # Baked into the add-in manifest's URLs, so it must be the port actually
    # served rather than whatever the default happens to be.
    bind_port = port if port is not None else listener_port()

    def _authorized() -> bool:
        header = str(request.headers.get("Authorization", ""))
        prefix = "Bearer "
        if not header.startswith(prefix):
            return False
        # Constant-time compare so the token can't be recovered by timing.
        return hmac.compare_digest(header[len(prefix) :].strip(), expected_token)

    @app.get("/api/health")
    def health() -> Response:
        # Unauthenticated on purpose: the task pane uses this to tell "the
        # desktop app isn't running" apart from "the token is wrong", and it
        # reveals nothing beyond the fact that the listener is up.
        return jsonify({"status": "ok", "service": "odd_job_intake"})

    @app.get("/job-intake-root-ca.crt")
    def root_ca() -> Response:
        _cert, _key, ca_path = ensure_loopback_certificate()
        return Response(
            ca_path.read_bytes(),
            mimetype="application/x-x509-ca-cert",
            headers={"Content-Disposition": "attachment; filename=job-intake-root-ca.crt"},
        )

    # --- Outlook add-in (Phase 3) --------------------------------------------
    # These are unauthenticated because Office fetches them before any add-in
    # code runs and cannot attach a bearer token. That is safe here: they are
    # reachable only from this machine over loopback, and the only sensitive
    # value among them - the API token - is injected into the rendered pane,
    # which is itself only reachable from this machine.

    def _addin_base_url() -> str:
        host = str(request.args.get("host", "") or "").strip().casefold()
        if host not in ALLOWED_ADDIN_HOSTS:
            host = DEFAULT_ADDIN_HOST
        return f"https://{host}:{bind_port}"

    @app.get("/addin/manifest.xml")
    def addin_manifest() -> Response:
        """The sideloadable manifest, rendered with the live port.

        Generated rather than stored so its URLs can never drift out of sync
        with the port the listener is actually on - a mismatch shows up in
        Outlook as an add-in that simply refuses to load. Pass ?host=127.0.0.1
        to switch loopback spelling; the certificate covers both.
        """
        xml = render_template("job_intake_manifest.xml", base_url=_addin_base_url())
        return Response(
            xml,
            mimetype="application/xml",
            headers={"Content-Disposition": "attachment; filename=job-intake-manifest.xml"},
        )

    @app.get("/addin/taskpane")
    def addin_taskpane() -> Response:
        html = render_template(
            "job_intake_taskpane.html",
            base_url=_addin_base_url(),
            api_token=expected_token,
        )
        # The pane embeds the API token, so it must never be cached to disk by
        # the webview.
        return Response(
            html,
            mimetype="text/html",
            headers={"Cache-Control": "no-store, no-cache, must-revalidate", "Pragma": "no-cache"},
        )

    @app.get("/addin/static/<path:filename>")
    def addin_static(filename: str) -> Response:
        # send_from_directory rejects traversal outside the directory itself.
        return send_from_directory(ADDIN_STATIC_DIR, filename)

    @app.get("/api/job-intake/check")
    def check_job_number() -> tuple[Response, int] | Response:
        if not _authorized():
            return jsonify({"error": "Unauthorized."}), 401
        job_number = str(request.args.get("job_number", "")).strip().upper()
        try:
            job_intake_service.resolve_job_root(job_number)
            exists = job_intake_service.job_folder_exists(job_number)
            label_required = job_intake_service.label_required_for(job_number)
            placeholder = job_intake_service.is_placeholder_job_number(job_number)
        except JobIntakeError as exc:
            return jsonify({"error": str(exc)}), 400
        # Drives whether the caller asks for a Label: an existing job folder
        # means this one-off must nest under one, and so does a placeholder,
        # where several jobs would otherwise share a folder while they wait
        # for their real numbers.
        return jsonify(
            {
                "job_number": job_number,
                "exists": exists,
                "label_required": label_required,
                "placeholder": placeholder,
            }
        )

    @app.get("/api/job-intake/status")
    def intake_status() -> tuple[Response, int] | Response:
        """Whether a queued intake has finished, and its summary once it has.

        Polled by the caller after a submit: the slow half runs on a background
        thread and there is no way to push the result back, so this is how a
        summary reaches anything outside this process.
        """
        if not _authorized():
            return jsonify({"error": "Unauthorized."}), 401
        key = str(request.args.get("key", "")).strip()
        entry = job_intake_registry.get_entry(key)
        if entry is None:
            return jsonify({"error": f"No intake found for {key}."}), 404

        state = job_intake_registry.entry_state(entry)
        # `done` is what a poller waits on, and it is true for a failure too -
        # otherwise a job that has already failed keeps the caller waiting out
        # its full timeout. `complete` stays as-is for older callers, meaning
        # specifically "finished successfully".
        complete = bool(entry.get("complete"))
        return jsonify(
            {
                "key": key,
                "complete": complete,
                "state": state,
                "done": state in job_intake_registry.TERMINAL_STATES,
                "status": entry.get("status"),
                "error": entry.get("error"),
                "parts": len(entry.get("material_qty", [])),
                "po_number": entry.get("po_number"),
                "due_date": entry.get("due_date"),
                "job_folder": entry.get("job_folder"),
                # Composed server-side so the wording can change without
                # touching the macro.
                "summary": job_intake_service.intake_summary_text(entry) if complete else "",
            }
        )

    @app.post("/api/job-intake")
    def submit_job_intake() -> tuple[Response, int] | Response:
        if not _authorized():
            return jsonify({"error": "Unauthorized."}), 401
        payload = request.get_json(silent=True)
        if not isinstance(payload, dict):
            return jsonify({"error": "Send a JSON object."}), 400

        job_number = str(payload.get("job_number", "")).strip().upper()
        label = str(payload.get("label", "") or "").strip()
        email_body = str(payload.get("email_body", "") or "")
        # Not a context manager: the slow half runs on a background thread and
        # would find the directory gone. It is removed when that finishes.
        staging = tempfile.mkdtemp(prefix="odd_job_intake_")
        try:
            if True:
                attachments = payload.get("attachments") or []
                files = (
                    _decode_attachments(attachments, Path(staging)) if attachments else []
                )
                # A job can arrive as a path to W: with no attachments at all,
                # so "nothing was attached" is only fatal when the body names
                # no folder either.
                if not any(path.name.casefold().endswith(".dxf") for path in files):
                    if not job_intake_service.paths_in_text(email_body):
                        return (
                            jsonify(
                                {
                                    "error": (
                                        "No .dxf attachment and no folder path in the "
                                        "email body - there is nothing to nest."
                                    )
                                }
                            ),
                            400,
                        )

                # Only the fast, decisive half runs here: resolve the paths,
                # enforce the fresh-vs-label rule, make the folder. That is
                # what makes a second click fail on the already-exists guard
                # rather than racing the first.
                entry, _paths = job_intake_service.begin_intake(
                    job_number,
                    label or None,
                    source="outlook",
                    email_subject=str(payload.get("email_subject", "") or ""),
                    email_sender=str(payload.get("email_sender", "") or ""),
                )
        except JobIntakeError as exc:
            shutil.rmtree(staging, ignore_errors=True)
            return jsonify({"error": str(exc)}), 400
        except ValueError as exc:
            # Raised by the registry when this job/label was already filed.
            shutil.rmtree(staging, ignore_errors=True)
            return jsonify({"error": str(exc)}), 409
        except Exception:
            shutil.rmtree(staging, ignore_errors=True)
            _logger.exception("Job intake from Outlook failed for %r.", job_number)
            return jsonify({"error": "The intake failed; check the desktop app's log."}), 500

        # The slow half - copying files off W:, scraping the PDFs, converting a
        # BOM - runs on a background thread. A real job took 53s where the
        # caller's socket gives up at 30, and it reported a success as a
        # failure. There is no way to push the result back afterwards, so the
        # answers that need the work done are returned as "pending" and the
        # desktop queue fills them in as it polls.
        if run_in_background:
            threading.Thread(
                target=_finish_intake,
                args=(entry, files, email_body, staging),
                name=f"job-intake-{entry.get('key')}",
                daemon=True,
            ).start()
        else:
            _finish_intake(entry, files, email_body, staging)

        _logger.info("Queued Outlook intake %s (%d attachments).", entry.get("key"), len(files))
        return (
            jsonify(
                {
                    "key": entry.get("key"),
                    "job_number": entry.get("job_number"),
                    "label": entry.get("label"),
                    # Known only after the files are read - the existing macro
                    # prints whatever it finds here, so these read as "pending"
                    # rather than as a blank or a wrong zero.
                    "po_number": entry.get("po_number") if not run_in_background else "pending",
                    "due_date": entry.get("due_date") if not run_in_background else "pending",
                    "parts": (
                        len(entry.get("material_qty", []))
                        if not run_in_background
                        else "pending"
                    ),
                    # Derived from the job number alone, so it is real already.
                    "job_folder": entry.get("job_folder"),
                    "status": entry.get("status") if not run_in_background else "queued",
                    # Where to look for the rest: anything non-terminal means
                    # the caller should poll /api/job-intake/status.
                    "state": (
                        job_intake_registry.entry_state(
                            job_intake_registry.get_entry(str(entry.get("key", "")))
                        )
                        if not run_in_background
                        else job_intake_registry.STATE_QUEUED
                    ),
                    "attachments": (
                        [str(item.get("filename")) for item in entry.get("attachments", [])]
                        if not run_in_background
                        else [Path(path).name for path in files]
                    ),
                    "po_unmatched": entry.get("po_unmatched", []) if not run_in_background else [],
                }
            ),
            201,
        )

    return app


def _ssl_context():
    import ssl

    cert_path, key_path, _ca_path = ensure_loopback_certificate()
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=str(cert_path), keyfile=str(key_path))
    return context


def serve_tls(port: int | None = None, *, token: str | None = None) -> None:
    """Serve over HTTPS on 127.0.0.1 until the process exits.

    Uses Flask/werkzeug's WSGI server with a TLS context rather than waitress:
    waitress has no built-in TLS termination, and this listener handles one
    user's occasional single-request submissions, so a threaded WSGI server is
    the right size for the job.
    """
    bind_port = port if port is not None else listener_port()
    app = create_app(token, port=bind_port)
    app.run(
        host=HOST,
        port=bind_port,
        ssl_context=_ssl_context(),
        threaded=True,
        debug=False,
        use_reloader=False,
    )


def start_listener(port: int | None = None) -> threading.Thread:
    """Start the listener on a daemon thread and return it.

    Daemon so it dies with the desktop app - matching the accepted tradeoff
    that Outlook intake only works while the app is open, on this PC.
    Failures are logged, never raised: the desktop app must still start if
    the port is taken or the cert can't be written.
    """
    bind_port = port if port is not None else listener_port()

    def _run() -> None:
        try:
            serve_tls(bind_port)
        except Exception:
            _logger.exception("The job intake listener stopped on port %d.", bind_port)

    thread = threading.Thread(target=_run, name="job-intake-listener", daemon=True)
    thread.start()
    _logger.info("Job intake listener starting on https://%s:%d", HOST, bind_port)
    return thread


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    token_value = ensure_api_token()
    port_value = listener_port()
    _cert, _key, ca = ensure_loopback_certificate()
    print(f"Listening on https://{HOST}:{port_value}")
    print(f"Root CA to trust: {ca}")
    print(f"Bearer token: {token_value}")
    serve_tls(port_value)
