"""Standalone launcher for the Odd Job Intake page.

Runs the same JobIntakePage that master_app can embed, in its own window.
"""

from __future__ import annotations

import logging
import os
import sys


def _listener_enabled() -> bool:
    return os.environ.get("ODD_JOB_INTAKE_LISTENER", "1").strip().lower() not in {"0", "false", "no"}


def main() -> int:
    from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox

    from explorer_bridge import load_explorer_api
    from job_intake_page import JobIntakePage

    app = QApplication.instance() or QApplication(sys.argv)

    if _listener_enabled():
        # Daemon thread, loopback only. A failure here (port in use, no cert)
        # is logged inside start_listener and must never stop the desktop app.
        try:
            from job_intake_server import start_listener

            start_listener()
        except Exception:
            logging.getLogger(__name__).exception("Could not start the job intake listener.")

    try:
        explorer_api = load_explorer_api()
    except Exception as exc:  # surface the missing-sibling case instead of crashing
        QMessageBox.critical(None, "Odd Job Intake", str(exc))
        return 1

    page = JobIntakePage(explorer_api=explorer_api)
    window = QMainWindow()
    window.setWindowTitle("Odd Job Intake")
    window.setCentralWidget(page)
    window.resize(1280, 820)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
