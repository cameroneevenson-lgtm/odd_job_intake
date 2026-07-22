"""Standalone launcher for the Odd Job Intake page.

Runs the same JobIntakePage that master_app can embed, in its own window.
"""

from __future__ import annotations

import sys


def main() -> int:
    from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox

    from explorer_bridge import load_explorer_api
    from job_intake_page import JobIntakePage

    app = QApplication.instance() or QApplication(sys.argv)

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
