"""Structural checks on the Outlook macro.

There is no VBA compiler here, so this cannot prove JobIntake.bas builds. What
it does catch is the class of mistake that has actually bitten: a module-level
`Const` written between two procedures, which VBA rejects with "Only comments
may appear after End Sub/End Function". That went out twice, each time costing
a trip to the machine to find out. It is cheap to check from here and it is not
otherwise checked by anything.
"""

from __future__ import annotations

import json
from pathlib import Path
import re

import pytest

from paths import APP_DIR
from tests import vba_json


BAS_PATH = APP_DIR / "outlook_vba" / "JobIntake.bas"

# A line that opens or closes a procedure. VBA scoping keywords are optional.
_PROC_START = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property)\b",
    re.IGNORECASE,
)
_PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
# Module-level declarations, which must all sit above the first procedure.
_DECLARATION = re.compile(
    r"^\s*(?:Public|Private|Global|Dim)\s+(?:Const\s+|WithEvents\s+)?\w+\s+As\s+",
    re.IGNORECASE,
)
_CONST = re.compile(r"^\s*(?:Public\s+|Private\s+|Global\s+)?Const\b", re.IGNORECASE)


def _code_lines() -> list[tuple[int, str]]:
    """Numbered lines with comments and blanks dropped.

    Line continuations are joined so a wrapped statement is judged as the one
    statement it is.
    """
    raw = BAS_PATH.read_text(encoding="utf-8").splitlines()
    joined: list[tuple[int, str]] = []
    pending = ""
    pending_number = 0
    for number, line in enumerate(raw, start=1):
        stripped = line.strip()
        if not pending and (not stripped or stripped.startswith("'")):
            continue
        if not pending:
            pending_number = number
        pending += stripped
        if pending.endswith("_"):
            pending = pending[:-1] + " "
            continue
        joined.append((pending_number, pending))
        pending = ""
    if pending:
        joined.append((pending_number, pending))
    return joined


def test_the_macro_file_is_present_and_named_for_import() -> None:
    assert BAS_PATH.exists(), f"{BAS_PATH} is what the user imports into Outlook"
    first = BAS_PATH.read_text(encoding="utf-8").splitlines()[0]
    # Outlook takes the module name from this attribute, not the filename.
    assert first.strip() == 'Attribute VB_Name = "JobIntake"'
    assert "Option Explicit" in BAS_PATH.read_text(encoding="utf-8")


def test_every_procedure_is_closed_exactly_once() -> None:
    depth = 0
    for number, line in _code_lines():
        if _PROC_END.match(line):
            depth -= 1
            assert depth >= 0, f"line {number}: End without a matching opener"
        elif _PROC_START.match(line):
            depth += 1
            assert depth == 1, f"line {number}: a procedure opened inside another"
    assert depth == 0, "a procedure was left unclosed"


def test_no_module_level_declaration_sits_between_procedures() -> None:
    """The compile error that has gone out twice.

    VBA allows module-level `Const`/`Dim` only in the declarations section above
    the first procedure. One placed after an `End Sub` fails to compile with a
    message that names the *procedure*, not the declaration, so it reads as if
    the procedure is at fault.
    """
    inside_procedure = False
    seen_first_procedure = False
    offenders: list[str] = []

    for number, line in _code_lines():
        if _PROC_END.match(line):
            inside_procedure = False
            continue
        if _PROC_START.match(line):
            inside_procedure = True
            seen_first_procedure = True
            continue
        if inside_procedure or not seen_first_procedure:
            continue
        if _CONST.match(line) or _DECLARATION.match(line):
            offenders.append(f"line {number}: {line}")

    assert not offenders, (
        "module-level declarations must live above the first procedure; VBA "
        "reports these as 'Only comments may appear after End Sub':\n  "
        + "\n  ".join(offenders)
    )


@pytest.mark.parametrize(
    "name",
    ["SendToJobIntake", "JsonValue", "HttpCall", "WaitForIntakeAndReply", "DraftReply"],
)
def test_the_procedures_the_ribbon_and_flow_depend_on_still_exist(name: str) -> None:
    """SendToJobIntake is wired to a ribbon button by name - renaming it breaks
    the button silently, with no error until someone clicks it."""
    body = BAS_PATH.read_text(encoding="utf-8")
    assert re.search(rf"\b(?:Sub|Function)\s+{name}\b", body), f"{name} is gone"


def test_the_poll_loop_stops_on_a_failure_not_just_a_success() -> None:
    """A failed intake used to leave the macro polling to its full timeout and
    then report the job as "still working"."""
    body = BAS_PATH.read_text(encoding="utf-8")
    assert 'JsonValue(response, "done")' in body, (
        "the poll loop must exit on the server's terminal state, which covers "
        "a failure as well as a success"
    )


def test_the_macro_and_the_listener_agree_on_the_api_version() -> None:
    """The macro is imported by hand, so the two drift whenever either is
    changed alone. They are checked against each other here so that drift is
    caught in the repo rather than by a wrong field on the shop floor."""
    import job_intake_server

    declared = re.search(
        r"Private Const API_VERSION As Long = (\d+)", BAS_PATH.read_text(encoding="utf-8")
    )
    assert declared, "JobIntake.bas must declare API_VERSION"
    assert int(declared.group(1)) == job_intake_server.API_VERSION, (
        "JobIntake.bas and job_intake_server.API_VERSION disagree - bump both "
        "together, and have the user re-import the .bas"
    )


# --- the two routines that do real string work -------------------------------
#
# Ported to Python in vba_json.py so their logic can be exercised; see that
# module for why. These were wrong in ways that produced plausible output
# rather than an error.


@pytest.mark.parametrize(
    "raw",
    [
        "plain text",
        'he said "make 19"',
        r"W:\LASER\For Battleshield Fabrication\F59487",
        "line one\r\nline two\r\nW:\\LASER\\Job",
        "tab\there",
        "control\x01char",
        "quote \" and backslash \\ together",
        "",
    ],
    ids=[
        "plain", "quotes", "windows-path", "multiline", "tab", "control",
        "both-specials", "empty",
    ],
)
def test_what_the_macro_encodes_python_decodes_to_the_same_string(raw: str) -> None:
    """The direction that carries the email body.

    Python's json.loads is the real reader on the other end, so it is the
    authority here rather than a second hand-rolled decoder.
    """
    payload = '{"email_body":"' + vba_json.json_escape(raw) + '"}'
    assert json.loads(payload)["email_body"] == raw


def test_line_breaks_survive_the_trip() -> None:
    """The specific bug. Newlines were replaced with spaces, which ran the body
    onto one line - and the server reads it line by line to find the W: path
    and a quantity, so both simply stopped being found."""
    body = "Please make these:\r\nW:\\LASER\\For Battleshield Fabrication\\F59487\r\nThanks"
    decoded = json.loads('{"b":"' + vba_json.json_escape(body) + '"}')["b"]

    assert decoded == body
    lines = decoded.splitlines()
    assert len(lines) == 3
    assert lines[1] == r"W:\LASER\For Battleshield Fabrication\F59487"


@pytest.mark.parametrize(
    "value",
    [
        "plain",
        'the print says "11GA" but the BOM says 0.12',
        r"L:\BATTLESHIELD\F-LARGE FLEET\F59487",
        "first line\nsecond line",
        "trailing backslash \\",
        "unicode é and ’",
        "",
    ],
    ids=["plain", "quotes", "path", "multiline", "backslash", "unicode", "empty"],
)
def test_the_macro_reads_back_exactly_what_the_server_sent(value: str) -> None:
    """The direction that carries error messages and folder paths.

    The old reader stopped at the first quote - truncating any message that
    contained one - and "unescaped" by replacing \\\\" and \\\\\\\\, sequences
    JSON never emits, so real escapes came through untouched.
    """
    payload = json.dumps({"before": 1, "error": value, "after": "x"})
    assert vba_json.json_value(payload, "error") == value


def test_the_reader_still_handles_bare_literals_and_missing_keys() -> None:
    payload = json.dumps(
        {"complete": True, "done": False, "parts": 13, "error": None, "key": "F59487"}
    )
    assert vba_json.json_value(payload, "complete") == "true"
    assert vba_json.json_value(payload, "done") == "false"
    assert vba_json.json_value(payload, "parts") == "13"
    assert vba_json.json_value(payload, "error") == "null"
    assert vba_json.json_value(payload, "key") == "F59487"
    # Absent keys come back empty, which is what the macro checks with Len().
    assert vba_json.json_value(payload, "nope") == ""


def test_a_real_status_response_reads_correctly_end_to_end() -> None:
    """Against a response the listener actually produces, quotes and all."""
    payload = json.dumps(
        {
            "key": "F59487",
            "complete": False,
            "state": "failed",
            "done": True,
            "status": "error",
            "error": 'Intake failed after queuing: BOM.csv uses "PLATE 11GA" with no rule',
            "parts": 0,
            "job_folder": r"L:\BATTLESHIELD\F-LARGE FLEET\F59487",
            "summary": "",
        }
    )
    assert vba_json.json_value(payload, "done") == "true"
    assert vba_json.json_value(payload, "complete") == "false"
    assert vba_json.json_value(payload, "job_folder") == r"L:\BATTLESHIELD\F-LARGE FLEET\F59487"
    # The whole message, not the part before the first quote.
    assert vba_json.json_value(payload, "error").endswith('with no rule')
    assert '"PLATE 11GA"' in vba_json.json_value(payload, "error")
