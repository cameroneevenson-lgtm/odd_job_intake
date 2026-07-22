# CLAUDE.md — odd_job_intake

Intake + setup for **one-off fabrication jobs** (`M`/`W`/`S`/`F`/`P` + digits)
that skip the truck/kit scaffold the rest of the shop tooling assumes, but
still want a job folder, a blank RADAN project, parts registered with
material/qty, and block files sent to the machine.

Flat layout (no packages), matching the other `C:\Tools` shop apps. Run tests
and the app from the repo root with the shared venv:

```
C:\Tools\.venv\Scripts\python.exe -m pytest -q
C:\Tools\.venv\Scripts\python.exe app.py
```

## Module roles

| File | Role |
| --- | --- |
| `paths.py` | Shop paths, `JOB_PREFIX_TO_ROOT`, registry location. Change shop locations only here. |
| `job_intake_registry.py` | Atomic JSON registry of intakes (`_runtime/job_intake_registry.json`). No Qt. |
| `job_intake_service.py` | All orchestration: path resolution, folder creation, RPD clone, PO scraping, import-CSV build, RADAN/block wrappers. **No Qt** — keep it that way so a future listener can reuse it. |
| `job_intake_page.py` | `JobIntakePage` Qt widget: queue table, parts grid, action buttons, background polling. |
| `explorer_bridge.py` | Lazily loads `truck_nest_explorer`'s RADAN import + block transfer. |
| `app.py` | Standalone window launcher. |
| `docs/PLAN.md` | Full design record incl. the unbuilt Outlook add-in phases. |

## Invariants — these are load-bearing, don't "simplify" them

- **PyMuPDF linearizes the PO table one _cell_ per line**, not one row per
  line: a filled row arrives as three consecutive lines (`"1"`, `"10"`,
  `"Clip-End - 1/4\" Mild Steel"`). The matcher is therefore
  *description-anchored* — it finds the description line, then reads the qty
  from the two integer lines above it. A row-based regex looks cleaner and
  **silently matches nothing**; this was found by testing against real POs.
- **Registry functions resolve `JOB_INTAKE_REGISTRY_PATH` at call time**
  (`_resolve_path`), not as a default argument. Default args bind at import
  time, which silently defeats monkeypatching and made tests write to the
  real registry.
- **`JobIntakePage` must not touch `explorer_api` during construction** —
  only inside button handlers. Tests (and any host embedding it) inject a
  stub `SimpleNamespace(services=...)`.
- **Importing `explorer_bridge` must never require `truck_nest_explorer`** —
  only calling `load_explorer_api()` does.
- **Never route a one-off job number through `truck_nest_explorer`'s
  truck-oriented path helpers** (`is_standard_truck_number`,
  `build_kit_paths`, `release_root_for_job`). This repo owns its own path
  resolution end to end, so an `F`/`P`-prefixed one-off can't be misrouted
  into the real truck machinery.
- **Job numbers are opaque strings** — prefix letter + digits, never
  validated to a fixed width (5 digits may become 6).

## Job folder convention (verified against real shop folders)

Prefix picks an existing `L:\BATTLESHIELD` root — **never invent a new root**.

```
Fresh job number:
  L:\BATTLESHIELD\<ROOT>\<job>\            <- DXFs + PO PDFs flat here
      <job>\                                <- inner project dir
          nests\  remnants\  <job>.rpd

Job number already exists (real truck, or a prior one-off) -> Label required:
  L:\BATTLESHIELD\<ROOT>\<job>\<label>\<job> <label>\<job> <label>.rpd
```

`create_job_folders` refuses a fresh intake when the folder already exists and
tells the user to add a Label — that guard is deliberate.

## Data-entry rules

- **Material stays manual.** Customer POs spell it inconsistently
  (`Aluminum`/`Aluminium`, `Stainless Steel 304`), so PO text is shown as
  read-only reference next to the field, never parsed into the dropdown.
- **Thickness is manual and required** (RADAN's import rejects `<= 0`); it is
  not derivable from material — the same material spans many thicknesses.
- **Strategy is derived from material** (Aluminum→`Air`, Mild Steel→`O2`,
  Stainless→`N2`) and shown read-only.
- **Qty and due date may be pre-filled from the PO but stay editable** — the
  extraction is best-effort.
- Every PO line should have at least a partial DXF match; lines that match
  nothing surface in a warning banner (missing attachment, or an order-wide
  note like `ALL MATERIAL 1/8" MILD STEEL`).

## PO extraction reality (validated against real POs, 2022–2026)

Tolerate without erroring: `LASER ORDER` vs `LASER QUOTE`; `PO Number:` vs
`Purchase Order`; `Date Required` holding `RUSH`/`ASAP`/blank instead of a
date (→ `due_date=None`, `due_note` set); blank or typo'd Drawing Number
cells (`L1,DXF`, `End Cap D.dxf + End Cap D. Pdf`); bare part codes with no
material; multi-page POs. PO number matches `\d{4}-\d{3}` with optional
`PO-`/`Q-` prefix and `-R\d+` suffix. Match keys are punctuation-stripped and
casefolded, longest-stem-first (so `L1` doesn't steal `L1-CP`'s row).

A single date on the page is the *order* date, not a due date — only claim a
due date when two or more dates are present.

## Testing

Qt tests run offscreen (`QT_QPA_PLATFORM=offscreen`) with a module-scoped
`qapp` fixture. `tests/__init__.py` exists so pytest puts the repo root on
`sys.path`. PO tests build synthetic PDFs reproducing the verified
cell-per-line layout. Isolate filesystem/registry work by monkeypatching
`job_intake_service.BATTLESHIELD_ROOT`, `job_intake_service.EXPLORER_TEMPLATE_PATH`,
and `job_intake_registry.JOB_INTAKE_REGISTRY_PATH`.

## Status

Phase 1 (folder + RPD + parts grid + RADAN import + block transfer) is built
and tested. The **Outlook 365 add-in is designed but not built** — a local
`127.0.0.1` HTTPS listener that accepts a job number plus base64 DXF
attachments and calls the same service functions with `source="outlook"`,
plus a task-pane manifest. See `docs/PLAN.md` (Phases 2–3), including the
flagged risk that Office's webview may reject a self-signed cert.
