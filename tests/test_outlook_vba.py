"""Structural checks on the Outlook macro.

There is no VBA compiler here, so this cannot prove JobIntake.bas builds. What
it does catch is the class of mistake that has actually bitten: a module-level
`Const` written between two procedures, which VBA rejects with "Only comments
may appear after End Sub/End Function". That went out twice, each time costing
a trip to the machine to find out. It is cheap to check from here and it is not
otherwise checked by anything.
"""

from __future__ import annotations

from pathlib import Path
import re

import pytest

from paths import APP_DIR


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
