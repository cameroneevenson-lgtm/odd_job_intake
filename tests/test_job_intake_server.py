"""Tests for the loopback listener the Outlook add-in posts to.

These drive the Flask app through its test client rather than binding a real
socket; the real HTTPS bind is exercised by the live verification described in
docs/PLAN.md. The point of these is that the listener delegates to
job_intake_service instead of growing its own intake logic, and that it
rejects the untrusted parts of a request.
"""

from __future__ import annotations

import base64
from pathlib import Path

import pytest

import job_intake_registry
import job_intake_server
import job_intake_service


TOKEN = "test-token-value"
AUTH = {"Authorization": f"Bearer {TOKEN}"}


@pytest.fixture
def client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """A listener wired to a throwaway L: root and registry."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    app = job_intake_server.create_app(token=TOKEN)
    app.config["TESTING"] = True
    return app.test_client()


def _attachment(name: str, content: bytes = b"0\nSECTION\n") -> dict[str, str]:
    return {"name": name, "contentBytes": base64.b64encode(content).decode("ascii")}


def _payload(**overrides):
    payload = {
        "job_number": "M90001",
        "attachments": [_attachment("Clip-End.dxf")],
        "email_subject": "Laser order",
        "email_sender": "buyer@example.com",
    }
    payload.update(overrides)
    return payload


# --- auth --------------------------------------------------------------------


def test_health_needs_no_token(client) -> None:
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.get_json()["status"] == "ok"


@pytest.mark.parametrize(
    "headers",
    [{}, {"Authorization": "Bearer wrong"}, {"Authorization": TOKEN}],
    ids=["missing", "wrong-token", "no-bearer-prefix"],
)
def test_submit_rejects_bad_authorization(client, headers) -> None:
    response = client.post("/api/job-intake", json=_payload(), headers=headers)
    assert response.status_code == 401


# --- the happy path delegates to the service ---------------------------------


def test_submit_creates_folder_and_registry_entry_with_outlook_source(client) -> None:
    response = client.post("/api/job-intake", json=_payload(), headers=AUTH)
    assert response.status_code == 201, response.get_json()
    body = response.get_json()

    assert body["job_number"] == "M90001"
    assert body["parts"] == 1
    assert body["attachments"] == ["Clip-End.dxf"]

    # The DXF really landed in the shop's folder shape, not just in the JSON.
    job_folder = Path(body["job_folder"])
    assert job_folder.parts[-2:] == ("M-FABRICATION", "M90001")
    assert (job_folder / "Clip-End.dxf").exists()

    entries = job_intake_registry.load_entries()
    assert len(entries) == 1
    entry = entries[0]
    assert entry["source"] == "outlook"
    assert entry["status"] == job_intake_registry.STATUS_NEW
    assert entry["email_sender"] == "buyer@example.com"
    assert entry["email_subject"] == "Laser order"
    # One seeded part row per DXF, awaiting manual material/thickness.
    assert entry["material_qty"] == [
        {
            "filename": "Clip-End.dxf",
            "material": "",
            "thickness": 0.0,
            "qty": 1,
            "unit": "in",
            "strategy": "",
            "po_ref": "",
        }
    ]


def test_submit_does_not_create_an_rpd_or_touch_radan(client) -> None:
    """RADAN work stays a desktop action - the listener must only stage files."""
    response = client.post("/api/job-intake", json=_payload(), headers=AUTH)
    assert response.status_code == 201

    job_folder = Path(response.get_json()["job_folder"])
    assert not list(job_folder.rglob("*.rpd"))
    assert not list(job_folder.rglob("*.csv"))
    assert response.get_json()["status"] == job_intake_registry.STATUS_NEW


def test_label_nests_the_job_under_its_own_subfolder(client, tmp_path: Path) -> None:
    (tmp_path / "L" / "M-FABRICATION" / "M90002").mkdir(parents=True)

    response = client.post(
        "/api/job-intake",
        json=_payload(job_number="M90002", label="Rework"),
        headers=AUTH,
    )
    assert response.status_code == 201, response.get_json()
    job_folder = Path(response.get_json()["job_folder"])
    assert job_folder.name == "Rework"
    assert (job_folder / "Clip-End.dxf").exists()


def test_existing_job_number_without_a_label_is_refused(client, tmp_path: Path) -> None:
    (tmp_path / "L" / "M-FABRICATION" / "M90003").mkdir(parents=True)

    response = client.post("/api/job-intake", json=_payload(job_number="M90003"), headers=AUTH)
    assert response.status_code == 400
    assert "Label" in response.get_json()["error"]


def test_refiling_the_same_job_asks_for_a_label(client) -> None:
    """The folder guard fires before the registry's duplicate-key check, so a
    resubmit is reported as "this job already has a folder, give it a Label"
    rather than a bare conflict - the same guidance the desktop app gives."""
    assert client.post("/api/job-intake", json=_payload(), headers=AUTH).status_code == 201

    second = client.post("/api/job-intake", json=_payload(), headers=AUTH)
    assert second.status_code == 400
    assert "Label" in second.get_json()["error"]


def test_a_registry_entry_without_its_folder_conflicts(client) -> None:
    """Reaches the 409 branch: the job folder was removed by hand on L: but
    the intake is still registered, so the append is what fails."""
    import shutil

    first = client.post("/api/job-intake", json=_payload(), headers=AUTH)
    assert first.status_code == 201
    shutil.rmtree(Path(first.get_json()["job_folder"]))

    second = client.post("/api/job-intake", json=_payload(), headers=AUTH)
    assert second.status_code == 409
    assert "M90001" in second.get_json()["error"]


# --- untrusted request handling ----------------------------------------------


def test_attachment_names_cannot_escape_the_job_folder(client, tmp_path: Path) -> None:
    response = client.post(
        "/api/job-intake",
        json=_payload(attachments=[_attachment(r"..\..\evil.dxf")]),
        headers=AUTH,
    )
    assert response.status_code == 201, response.get_json()

    # The traversal is stripped to a bare filename, so nothing is written
    # outside the job folder.
    job_folder = Path(response.get_json()["job_folder"])
    assert (job_folder / "evil.dxf").exists()
    assert not (tmp_path / "evil.dxf").exists()
    assert not (tmp_path / "L" / "evil.dxf").exists()


def test_non_dxf_pdf_attachments_are_refused(client) -> None:
    response = client.post(
        "/api/job-intake",
        json=_payload(attachments=[_attachment("Clip-End.dxf"), _attachment("logo.png")]),
        headers=AUTH,
    )
    assert response.status_code == 400
    assert "logo.png" in response.get_json()["error"]


def test_a_request_with_no_dxf_is_refused(client) -> None:
    response = client.post(
        "/api/job-intake",
        json=_payload(attachments=[_attachment("PO-8497-005.pdf", b"%PDF-1.4\n")]),
        headers=AUTH,
    )
    assert response.status_code == 400
    assert "dxf" in response.get_json()["error"].casefold()


@pytest.mark.parametrize(
    "attachments, expected",
    [
        ([], "at least one"),
        ([{"name": "a.dxf", "contentBytes": "not base64!!"}], "base64"),
        ([{"name": "a.dxf", "contentBytes": ""}], "empty"),
        ([{"name": "", "contentBytes": "AAAA"}], "unusable"),
    ],
    ids=["empty-list", "bad-base64", "empty-content", "blank-name"],
)
def test_malformed_attachments_are_refused(client, attachments, expected) -> None:
    response = client.post(
        "/api/job-intake", json=_payload(attachments=attachments), headers=AUTH
    )
    assert response.status_code == 400
    assert expected in response.get_json()["error"].casefold()


def test_bad_job_number_is_refused_before_anything_is_written(client, tmp_path: Path) -> None:
    response = client.post("/api/job-intake", json=_payload(job_number="X123"), headers=AUTH)
    assert response.status_code == 400
    assert not (tmp_path / "L").exists()


def test_non_json_body_is_refused(client) -> None:
    response = client.post("/api/job-intake", data="not json", headers=AUTH)
    assert response.status_code == 400


# --- the check route the task pane uses to decide about the Label field ------


def test_check_reports_whether_a_label_is_required(client, tmp_path: Path) -> None:
    fresh = client.get("/api/job-intake/check?job_number=M90004", headers=AUTH)
    assert fresh.status_code == 200
    assert fresh.get_json()["label_required"] is False

    (tmp_path / "L" / "M-FABRICATION" / "M90005").mkdir(parents=True)
    existing = client.get("/api/job-intake/check?job_number=m90005", headers=AUTH)
    assert existing.status_code == 200
    assert existing.get_json() == {
        "job_number": "M90005",
        "exists": True,
        "label_required": True,
    }


def test_check_rejects_an_unknown_prefix(client) -> None:
    response = client.get("/api/job-intake/check?job_number=Z12345", headers=AUTH)
    assert response.status_code == 400


def test_check_needs_a_token(client) -> None:
    assert client.get("/api/job-intake/check?job_number=M90006").status_code == 401


# --- the Outlook add-in surface ----------------------------------------------


@pytest.fixture
def addin_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """A listener on a known port, so manifest URLs are predictable."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    app = job_intake_server.create_app(token=TOKEN, port=9999)
    app.config["TESTING"] = True
    return app.test_client()


def test_manifest_is_well_formed_and_points_at_the_live_port(addin_client) -> None:
    """A manifest whose URLs don't match the served port produces an add-in
    that silently refuses to load, so the port is generated, never hardcoded."""
    from xml.etree import ElementTree

    response = addin_client.get("/addin/manifest.xml")
    assert response.status_code == 200

    xml = response.get_data(as_text=True)
    ElementTree.fromstring(xml)  # raises if the manifest isn't valid XML

    assert "https://localhost:9999/addin/taskpane" in xml
    assert "https://localhost:9999/addin/static/icon-16.png" in xml
    assert ":8790" not in xml


def test_manifest_can_switch_loopback_spelling(addin_client) -> None:
    """The cert covers both spellings; ?host= picks which one Office registers."""
    xml = addin_client.get("/addin/manifest.xml?host=127.0.0.1").get_data(as_text=True)
    assert "https://127.0.0.1:9999/addin/taskpane" in xml
    assert "https://localhost:9999" not in xml


def test_manifest_ignores_an_unrecognized_host(addin_client) -> None:
    """?host= must not become a way to point the add-in at an arbitrary origin."""
    xml = addin_client.get("/addin/manifest.xml?host=evil.example.com").get_data(as_text=True)
    assert "evil.example.com" not in xml
    assert "https://localhost:9999/addin/taskpane" in xml


def test_manifest_declares_the_requirement_set_the_pane_actually_needs(addin_client) -> None:
    """getAttachmentContentAsync is Mailbox 1.8; declaring lower would let the
    add-in install somewhere it then fails at runtime."""
    xml = addin_client.get("/addin/manifest.xml").get_data(as_text=True)
    assert 'DefaultMinVersion="1.8"' in xml
    assert "<Permissions>ReadItem</Permissions>" in xml
    assert "MessageReadCommandSurface" in xml


def test_taskpane_injects_the_token_and_is_not_cached(addin_client) -> None:
    response = addin_client.get("/addin/taskpane")
    assert response.status_code == 200

    html = response.get_data(as_text=True)
    # Injected server-side so it is never hardcoded in the JS or typed by hand.
    assert TOKEN in html
    assert "https://localhost:9999" in html
    # It embeds a credential, so the webview must not persist it.
    assert "no-store" in response.headers.get("Cache-Control", "")


def test_addin_static_files_are_served(addin_client) -> None:
    for name in ("taskpane.js", "taskpane.css", "icon-16.png", "icon-80.png"):
        response = addin_client.get(f"/addin/static/{name}")
        assert response.status_code == 200, name
        assert response.get_data()


def test_addin_static_refuses_traversal(addin_client) -> None:
    response = addin_client.get("/addin/static/../../paths.py")
    assert response.status_code in (403, 404)


def test_addin_routes_need_no_token(addin_client) -> None:
    """Office fetches these before any add-in code runs and cannot present a
    token; they're safe because they're reachable only over loopback."""
    for path in ("/addin/manifest.xml", "/addin/taskpane", "/addin/static/taskpane.js"):
        assert addin_client.get(path).status_code == 200, path
