# odd_job_intake

Intake and setup for one-off fabrication jobs — the quick `M`/`W`/`S`/`F`/`P`
jobs that don't fit the truck/kit structure the rest of the shop tooling
assumes, but still want the same conveniences: a job folder, a blank RADAN
project, parts registered with material/quantity, and block files sent to the
machine.

## What it does

Given a job number and a set of DXF files (plus an optional customer PO PDF),
the app:

1. Creates the job folder under the right `L:\BATTLESHIELD` root — chosen by
   the job number's prefix letter, matching the shop's existing layout (e.g.
   `M59919` → `L:\BATTLESHIELD\M-FABRICATION\M59919\`). A job number that
   already has a folder requires a **Label** so the one-off nests in its own
   subfolder instead of mixing into the existing job.
2. Scrapes the PO PDF (best-effort) for the PO number, due date, and a
   per-DXF quantity, and flags any PO line that has no matching DXF.
3. Lets you set material / thickness / quantity per DXF (strategy is derived
   from the material), plus a due date and Laser / Bend time estimates.
4. Clones the blank RADAN template into `<job>/<job>.rpd`.
5. Runs the headless RADAN import to convert the DXFs and register the parts.
6. After you nest in RADAN, sends the block files to the machine (verified copy).

Every intake is tracked in a small JSON registry (`_runtime/`), so the queue
survives restarts and can later be fed by an Outlook add-in as well as the
manual button.

## Running

```
C:\Tools\.venv\Scripts\python.exe app.py
```

The RADAN import and block-transfer steps borrow behavior from the sibling
`truck_nest_explorer` app at runtime (see `explorer_bridge.py`); those steps
need it installed under `C:\Tools`. Everything else (folder/RPD creation, PO
parsing, the parts grid) runs on its own.

`master_app` can also embed `JobIntakePage` as a tab by constructing it with
its own loaded explorer services — the page takes the services object as a
constructor argument and touches it only inside button handlers.

## Layout

| File | Role |
| --- | --- |
| `app.py` | Standalone window launcher |
| `job_intake_page.py` | `JobIntakePage` Qt widget (queue + parts grid + actions) |
| `job_intake_service.py` | Non-Qt orchestration: paths, RPD clone, PO extraction, CSV build, RADAN/block wrappers |
| `job_intake_registry.py` | Atomic JSON registry of intakes |
| `explorer_bridge.py` | Loads truck_nest_explorer's RADAN import + block transfer |
| `paths.py` | Shop paths and prefix→root mapping |
| `docs/PLAN.md` | Full design record, including the planned Outlook add-in phases |

## Tests

```
C:\Tools\.venv\Scripts\python.exe -m pytest -q
```

## Status

Phase 1 (the intake tab and its full flow) is complete. The Outlook 365
add-in — a local listener plus a task-pane button that hands an email's job
number and DXF attachments to this app — is designed but not yet built; see
`docs/PLAN.md`.
