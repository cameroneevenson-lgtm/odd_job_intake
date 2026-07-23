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
        # Safety checks before driving RADAN over COM. Both are the sibling
        # app's own, re-exported by its services module - a second
        # implementation here would be a second thing to get wrong about when
        # it is safe to touch RADAN.
        visible_radan_sessions=inventor_bridge.visible_radan_sessions,
        radan_csv_import_lock_status=inventor_bridge.radan_csv_import_lock_status,
    )
    return SimpleNamespace(services=services)


def load_import_log_dialog() -> type:
    """truck_nest_explorer's live import-log dialog.

    Reused rather than rebuilt so a RADAN import looks the same wherever it is
    started from, and so there is one place that knows how to tail that log.
    Imported only when an import actually runs - this module must stay
    importable without the sibling app.
    """
    root = str(TRUCK_NEST_EXPLORER_DIR.resolve())
    if not TRUCK_NEST_EXPLORER_DIR.exists():
        raise RuntimeError(f"truck_nest_explorer was not found at {root}")
    if root not in sys.path:
        sys.path.insert(0, root)
    module = importlib.import_module("dialogs.import_log_dialog")
    return module.ImportLogDialog
