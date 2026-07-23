"""Line-by-line ports of JobIntake.bas's JsonEscape and JsonValue.

There is no VBA interpreter here, and these two routines are the only place the
macro does real string work - they encode every email body sent to the listener
and decode every answer it gets back. Both were wrong in ways that produce
plausible-looking output rather than an error: newlines flattened to spaces, and
an "unescaper" looking for sequences JSON never emits.

A port is not the VBA, and it cannot catch a compile error - test_outlook_vba.py
does that. What it catches is the logic being wrong, which is what actually
shipped. Kept deliberately literal, in the same order and shape as the Basic, so
the two can be read side by side; that is worth more here than idiomatic Python.

If JobIntake.bas changes, change this with it.
"""

from __future__ import annotations


def json_escape(text: str) -> str:
    """Port of JsonEscape."""
    out: list[str] = []
    for ch in text:
        code = ord(ch)
        if code > 0xFFFF:
            # VBA works in UTF-16; a character above the BMP is a surrogate
            # pair there and passes through as-is either way.
            out.append(ch)
        elif code == 34:
            out.append('\\"')
        elif code == 92:
            out.append("\\\\")
        elif code == 8:
            out.append("\\b")
        elif code == 9:
            out.append("\\t")
        elif code == 10:
            out.append("\\n")
        elif code == 12:
            out.append("\\f")
        elif code == 13:
            out.append("\\r")
        elif code < 32:
            out.append("\\u" + format(code, "04X"))
        else:
            out.append(ch)
    return "".join(out)


def json_value(json_text: str, key: str) -> str:
    """Port of JsonValue.

    Returns "" when the key is absent, matching VBA's `Exit Function` leaving
    the return value at its empty-string default.
    """
    marker = f'"{key}":'
    start = json_text.find(marker)
    if start < 0:
        return ""
    i = start + len(marker)

    while i < len(json_text) and json_text[i] == " ":
        i += 1

    if i >= len(json_text) or json_text[i] != '"':
        # A bare literal: number, true, false, null.
        start = i
        while i < len(json_text) and json_text[i] not in ",}]":
            i += 1
        return json_text[start:i].strip()

    i += 1
    out: list[str] = []
    while i < len(json_text):
        ch = json_text[i]
        if ch == '"':
            break
        if ch == "\\":
            i += 1
            ch = json_text[i] if i < len(json_text) else ""
            if ch == "n":
                out.append("\n")
            elif ch == "r":
                out.append("\r")
            elif ch == "t":
                out.append("\t")
            elif ch == "b":
                out.append("\b")
            elif ch == "f":
                out.append("\f")
            elif ch == "u":
                out.append(chr(int(json_text[i + 1 : i + 5], 16)))
                i += 4
            else:
                # Covers \" \\ \/ - the character stands for itself.
                out.append(ch)
        else:
            out.append(ch)
        i += 1
    return "".join(out)
