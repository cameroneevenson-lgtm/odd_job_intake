"""Bridge to the sibling truck_nest_explorer app for the two RADAN operations
this feature needs: converting DXFs into an RPD's parts (headless import) and
copying nested block files to the machine.

JobIntakePage takes an ``explorer_api`` object with a ``.services`` attribute
exposing ``launch_radan_csv_import`` and ``send_project_block_files_to_machine``.
The desktop launcher builds that here; tests inject a stub instead, so importing
this module never requires truck_nest_explorer to be present - only calling
``load_explorer_api()`` does.
"""

from __future__ import annotations

import importlib
import sys
from types import SimpleNamespace

from paths import TRUCK_NEST_EXPLORER_DIR


def load_explorer_api() -> SimpleNamespace:
    root = str(TRUCK_NEST_EXPLORER_DIR.resolve())
    if not TRUCK_NEST_EXPLORER_DIR.exists():
        raise RuntimeError(
            f"truck_nest_explorer was not found at {root}; RADAN import and block "
            "transfer need it installed as a sibling app."
        )
    if root not in sys.path:
        sys.path.insert(0, root)
    inventor_bridge = importlib.import_module("inventor_bridge")
    w_block_transfer = importlib.import_module("w_block_transfer")
    services = SimpleNamespace(
        launch_radan_csv_import=inventor_bridge.launch_radan_csv_import,
        send_project_block_files_to_machine=w_block_transfer.send_project_block_files_to_machine,
    )
    return SimpleNamespace(services=services)
