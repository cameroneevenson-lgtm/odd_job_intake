# One-off job intake: Outlook -> master_app -> RADAN

## Context

The shop runs jobs that don't fit the canonical truck/kit-template structure `truck_nest_explorer` scaffolds (PAINT PACK, INTERIOR PACK, etc.) — quick fabrication/warranty/service/small-fleet jobs identified by a job number (prefix `M`/`W`/`S`/`F`/`P` + 5-6 digits) rather than a full truck build. These still deserve the same conveniences: a blank RPD, parts registered with material/qty, and "send blocks to machine" — without the kit scaffold or a row in `fabrication_flow_dashboard`.

The trigger should be an Outlook add-in action on a received email (attachments = DXFs): prompt for the job number, stage the attachments, then let the user assign material/thickness/qty per DXF (plus due date and Laser/Bend time estimates) inside a new `master_app` tab, which then generates the RPD and registers the parts via RADAN.

This plan was refined through direct discussion — several points below (folder convention, listener hosting, field names) are settled decisions from that conversation, not assumptions.

## Resolved design decisions

**Job folder convention — confirmed against real folders on `L:\BATTLESHIELD`, not invented.** The job number's prefix selects an existing top-level root (no new root folder):

```python
PREFIX_TO_ROOT = {
    "F": "F-LARGE FLEET", "P": "P-SMALL FLEET", "M": "M-FABRICATION",
    "W": "W-WARRANTY", "S": "S-SERVICE",
}
```

Confirmed live via `L:\BATTLESHIELD\M-FABRICATION\M59919\`: DXF/.sym files sit flat in the job folder alongside source PDFs, with an inner project folder repeating the job number:
```
L:\BATTLESHIELD\<ROOT>\<job_number>\
    <dxf files, flat>
    <job_number>\
        nests\
        remnants\
        <job_number>.rpd
```

**Two cases, auto-detected by checking whether `L:\BATTLESHIELD\<ROOT>\<job_number>\` already exists:**
- **Fresh job number** (no existing folder): create the shape above exactly.
- **Existing job number** (already a real truck, or a prior one-off): require a **Label** field from the user, and nest under it instead, mirroring the exact naming convention `truck_nest_explorer.services.build_kit_paths` already uses for kits/odd-jobs under a truck:
  ```
  L:\BATTLESHIELD\<ROOT>\<job_number>\<label>\<job_number> <label>\<job_number> <label>.rpd
  ```
  This keeps the one-off fully separated from whatever else already lives in that job's folder (kit subfolders, other one-offs). The Label field is only required in this branch.

Job numbers are treated as opaque strings (prefix in `{F,P,M,W,S}` + digits) — never validated to a fixed width (5 may become 6), and **never** passed through `truck_nest_explorer`'s truck-oriented path functions (`is_standard_truck_number`, `build_kit_paths`, `release_root_for_job`) — this feature owns its own path resolution end to end, so an F/P-prefixed one-off job number never risks being misrouted by the canonical truck machinery.

**Local listener hosting — embedded in master_app's own process, not `serve_web.py`.** Discussed at length: the separate Flask web companion (`serve_web.py`/`web_app.py`, autostarted via a Windows Scheduled Task) is explicitly on the back burner, and the user works from the desktop app only. Building on it would mean reviving a deprioritized subsystem for a "maybe someday" web future, when the actual reusable value (registry + service functions, see below) is UI/host-agnostic either way — moving to a web frontend later costs the same regardless of which option is chosen today. So: a small, separate listener runs *inside* `master_app`'s own desktop process (started alongside `OpsMainWindow`), bound to `127.0.0.1` only. It reuses `web_tls.ensure_lan_certificate()` directly (proven, already a plain importable function) purely to satisfy Office's HTTPS requirement — it does **not** import or depend on `web_app.py`, its session/login system, or `serve_web.py`'s autostart lifecycle at all. Tradeoff accepted: this only works while the desktop app is open, and only from the same PC — matching how the user actually works today.

**Material list**: no existing single source of truth. Read live from `inventor_to_radan/description_rules.csv`'s `Material` column (dedupe/casefold) with a hardcoded fallback list (`Aluminum 3003 CHK`, `Aluminum 5052`, `Mild Steel-A36`, `Stainless Steel`) if that CSV is missing/reshaped. `Aluminum 3003 CHK FTQ` is excluded from the user-facing dropdown — it's a forced per-part-number override elsewhere (`ftq_parts.csv`), not a real user choice.

**Strategy is auto-derived from Material** (not asked of the user), based on `description_rules.csv`'s observed 1:1 correlation: Aluminum -> `Air`, `Mild Steel-A36` -> `O2`, `Stainless Steel` -> `N2`. Shown read-only in the grid.

**Thickness is a required manual field per DXF** — not derivable from material (the same material spans many thicknesses in the real data) and hard-required by `import_parts_csv_headless.py`'s `_parse_thickness` (>0). This is a real scope addition beyond "material and qty" but unavoidable.

**Unit** defaults to `"in"` (matches `UNIT_TO_RADAN = {"mm": 0, "in": 1, "swg": 2}`), editable per row.

**Job-level fields** (not per-DXF): **Due Date**, **Laser** time estimate, **Bend** time estimate — using `fabrication_flow_dashboard/stages.py`'s own terminology (`Stage.LASER`/`Stage.BEND`, labels `"Laser"`/`"Bend"`), not "brake."

**PO attachment handling — confirmed against a real example (`M59919`'s own PO, "PFF Laser Order 8665-001 (Battleshield).pdf").** That job folder actually has *two* non-DXF PDFs (the PO itself, and a separate drawings PDF) — so there's no need to specially identify "which attachment is the PO"; the same best-effort extraction just runs against every non-DXF PDF attachment, and documents that don't match the expected shape simply contribute nothing.

The real PO's table has a **Drawing Number** column that is an exact, literal match to the DXF filename (`Clip-End.DXF` -> `Clip-End.DXF`), a clean numeric **Qty** column, and a **Date Required** field — all safe to auto-fill. Its **DESCRIPTION** column bundles the customer's own material wording and a fraction-notation thickness ("Clip-End - 1/4\" Mild Steel") in one string — genuinely ambiguous, left manual per the scoping discussion.

Design: `job_intake_service.extract_po_hints(po_path: Path, dxf_stems: list[str]) -> POHints` (new, uses `fitz`/PyMuPDF — add as a new master_app dependency; `truck_nest_explorer` already depends on it for a similar title-block text-extraction purpose, but this is a generic enough utility to own directly rather than reach through the embedding mechanism for it). `POHints` = `{po_number: str | None, due_date: date | None, line_items: dict[dxf_stem, {"qty": int | None, "raw_description": str | None}]}`. Heuristic, explicitly best-effort:
- Scan each page's text for a due-date label (`"Date Required"`, `"Due Date"`, etc.) followed by a parseable date.
- **PO number extraction**: every sampled real PO's number matches `\d{4}-\d{3}` (4 digits, hyphen, 3 digits), optionally prefixed (`PO-`, `Q-`) or suffixed with a revision (`-R\d+`) — e.g. `8497-005`, `7940-003`, `Q-7760-003`, `8645-008-R1`. Use `re.search(r"(?:PO-|Q-)?\d{4}-\d{3}(?:-R\d+)?", text)` near the "PO Number"/"Purchase Order" label as the `po_number` value, stored in the registry entry for the shop's own cross-referencing. The same pattern's presence anywhere on the page is also a useful secondary confidence signal that a given non-DXF attachment actually is a PO (alongside the table-structure signal from the Qty/Description columns), for jobs where multiple non-DXF PDFs are attached and only one is the real PO.

**Match key confirmed against a second real example (`M59841`'s PO, "PFF PO-8497-005") — the Drawing Number column is not reliable.** In `M59919` it was filled in with the exact DXF filename; in `M59841` it's blank for every line. The reliable match key across both is instead the **DESCRIPTION** column's leading part-name segment, which consistently follows a `"<Part Name> - <fraction>\" <Material word>"` shape (`"6X6 Cap - 1/8\" Aluminum"`, `"Base Plate - 70 Nicholas - 3/8\" Aluminum"`, `"Clip-End - 1/4\" Mild Steel"`) and matches the DXF stem case-insensitively. So: for each DXF's stem, search each DESCRIPTION line for that stem as a case-insensitive leading match (not a substring search anywhere in the page); if found, take that row's Qty column value; if the Drawing Number column is *also* present and matches, treat it as a confirming signal, but never require it. Capture the full DESCRIPTION line as `raw_description`.

(Incidentally, both examples' RPDs confirm the shop's thickness convention is the fraction rounded to 2 decimals via round-half-to-even — `1/8" -> 0.12`, `3/8" -> 0.38`, `1/4" -> 0.25` — exactly Python's default `round(x, 2)`. Not used now since Thickness stays manual per the reference-only decision, but worth knowing if thickness pre-fill is ever revisited later.)

**Real-world variance survey — 633 of these PO documents exist under `M-FABRICATION` alone (filename patterns include "PFF Laser Order *", "PFF PO-*", with/without "(Battleshield)", with "-R1" revision suffixes — filename is not a reliable filter, confirming the "run the heuristic against every non-DXF PDF attachment" design rather than matching by filename). Sampled 6 across a 4-year span and found real variance the extraction must tolerate without erroring:**
- **Template/label variants**: "LASER ORDER" vs. "LASER QUOTE"; field label "PO Number:" vs. "Purchase Order". The "Date Required" field can contain non-date text (`"RUSH"`) instead of an actual date — due-date parsing must fail silently (leave `due_date: None`), never raise.
- **DESCRIPTION format varies far more than the two-example pattern suggested**: sometimes the full `"<name> - <thickness>\" <material>"` shape; sometimes a bare code with no thickness/material at all (`"End Cap D"`, `"M1"`...`"M40"`); sometimes extra trailing descriptive suffixes after the material (`"L1-CP - 1/4\" Stainless Steel 304 - Checkered Plate"`); sometimes a **non-part summary row** giving shared material for the whole order (`"ALL MATERIAL 1/8\" MILD STEEL"`, not tied to any single line/DXF). Material spelling itself varies too (`"Aluminum"` vs. `"Aluminium"`, `"Stainless Steel 304"`) — further reinforcing that Material must stay manual/reference-only, never auto-parsed into the canonical dropdown.
- **Drawing Number column is unreliable in more ways than "sometimes blank"**: one real PO (41 line items, 2 pages) had it blank for every row; another combined DXF+PDF refs in one cell (`"End Cap D.dxf + End Cap D. Pdf"`, `"Channel 11a.dxf+.pdf"`); another had a literal typo (`"L1,DXF"` — comma instead of period). None of these are literal-equality matches against a real filename.
- **Match key must be normalized, not literal**: `"Channel 11a"` (DXF stem) vs. `"Channel - 11a"` (that PO's DESCRIPTION prefix) differ by an inserted `" - "`. The matching function must strip punctuation/collapse whitespace and casefold both sides before comparing (e.g. `re.sub(r"[^a-z0-9]", "", text.casefold())`), not do exact substring/prefix comparison.
- **Multi-page POs exist** (2 pages, 41 line items observed) — the extraction must scan every page's text, not just page 1.
- **Net effect**: for a meaningful fraction of real POs, zero line items will match anything (blank Drawing Number + bare unrelated codes with no matching DXF stems). That's fine and expected — it just means the reference panel and Qty/Due-Date pre-fill show nothing for that job, and the user falls back to fully manual entry, which is the baseline behavior anyway. This is not a bug to fix; it's the honest ceiling of a best-effort heuristic against inconsistent human-prepared documents.

In the Job Intake grid: **Qty and Due Date are pre-filled when a match is found, but remain fully editable** (never locked) since the extraction is heuristic. Material and Thickness are never auto-filled — instead, each row shows the matched `raw_description` (if any) as read-only reference text next to the Material/Thickness inputs, so the user can cross-check the customer's own wording while picking the canonical value themselves.

## Phase 1 — master_app Job Intake tab (manual trigger, RPD clone, RADAN import, send blocks)

**New files** (flat layout, matching master_app's existing convention):
- `master_app/job_intake_registry.py` — pure data-access module (no Qt, no Flask): atomic JSON load/save/append/update at `master_app/_runtime/job_intake_registry.json`, following `web_users.py`'s atomic-write pattern (`.tmp` + `Path.replace`). Schema per entry: `job_number, label (nullable), po_number (nullable, from PO extraction), received_at, updated_at, source ("manual"|"outlook"), status, email_subject, email_sender, job_folder, rpd_path, csv_log_path, error, due_date, laser_hours, bend_hours, attachments: [{filename, saved_path, size}], material_qty: [{filename, material, thickness, unit, qty, strategy}]`. Must be importable identically by both the desktop process and the Phase 2 listener (same file, shared schema).
- `master_app/job_intake_service.py` — non-Qt orchestration (mirrors the role of an existing `*_service.py` file): resolves `PREFIX_TO_ROOT`, detects fresh-vs-existing job folder, builds the correct path shape, copies DXFs in, `clone_rpd_template(...)` (new function modeled on `truck_nest_explorer/services.py`'s `_apply_template_project_defaults` literal-regex-substitution approach — reads `truck_nest_explorer/Template/Template.rpd`, substitutes job_number/label instead of truck_number/kit_name, rewrites `<JobName>`/`<NestFolder>`/`<RemnantSaveFolder>`, clones the template's `nests`/`remnants` subfolders), `build_import_csv_rows(...)` + `write_import_csv(...)` (exact 6-column, no-header format `read_import_csv` expects: `dxf_path, quantity, material, thickness, unit, strategy` — `dxf_path` must point at the copied-in path, not the original), `run_radan_import(...)`, `send_job_blocks_to_machine(...)`, `MATERIAL_DEFAULT_STRATEGY` lookup, material-list resolution.
- `master_app/ops_job_intake.py` — `JobIntakePage(QWidget)`, modeled directly on `ops_admin.py`'s `DiagnosticsPage` (header/refresh-button/content shape). Contains: a queue/list view of pending intakes (polling `QTimer` ~3-5s while visible + manual refresh button, matching the existing hot-reload timer precedent), a "Manual Intake" button (job number + Label-if-needed + file picker, for testing/use without Outlook), a form for Due Date / Laser hours / Bend hours, and a parts grid (one row per DXF) with delegates modeled on `radan_kitter/ui_parts_table.py`'s `KitComboDelegate`/`PrioritySpinDelegate` pattern: Material (combo), Thickness (new double-spin delegate), Qty (spin), Unit (combo), Strategy (read-only).

**Reused via the existing embedding mechanism** (`source_apps.load_truck_explorer_api()`, consistent with how the rest of master_app already reaches `truck_nest_explorer`/`radan_automation` — not a new direct-import pattern):
- `inventor_bridge.launch_radan_csv_import(...)` for the DXF->SYM conversion + part registration (real RADAN COM conversion, `--project-update-method=direct-xml`, explicitly **without** `--lab-symbol-writer`).
- `w_block_transfer.send_project_block_files_to_machine(...)` for "Send Blocks to Machine," with `release_root = L:\BATTLESHIELD\<matched ROOT>` and `machine_root` derived the same way the existing P-SMALL FLEET default is derived (sibling-name swap on `DEFAULT_MACHINE_EIA_ROOT`), `source_root` unchanged (shared `L:\BATTLESHIELD\BLOCK FILES`).

**Modified files:**
- `ops_main_window.py` — add `JobIntakePage` to the page tuple/stack loop (~line 237-242) and one sidebar nav button (~line 250-274 pattern).
- `ops_paths.py` — add `PREFIX_TO_ROOT`, `JOB_INTAKE_REGISTRY_PATH`, following existing constants style.

## Phase 2 — embedded local listener (in master_app's own process)

**New file:** `master_app/job_intake_server.py` — a small, separate Flask app (own instance, not `web_app.py`'s), served via `waitress` (already a dependency) on a background thread started from `app.py`'s `main()` or `OpsMainWindow.__init__`, bound to `127.0.0.1` on a new port (e.g. `8790`, configurable via a `MASTER_APP_JOB_INTAKE_PORT` env var following the existing `MASTER_WEB_*` naming convention). Reuses `web_tls.ensure_lan_certificate(...)` for the cert (own root CA download route mirroring `serve_web.py`'s `/master-ops-root-ca.crt`, so this listener is fully self-contained and never depends on `serve_web.py` running). Auth: a shared-secret bearer token, generated/stored the same way `web_users.py` handles its secret key (`master_app/_runtime/job_intake_api_token.key`, checked inline in the route — no session/login system at all, since this is a minimal single-purpose app).

Routes: `POST /api/job-intake` (job number + base64 attachments -> decode, filter to `.dxf`, delegate to `job_intake_service`/`job_intake_registry` — the *same* functions Phase 1 built, no duplicated logic) and, once Phase 3 exists, the task-pane HTML/JS/manifest routes (see below).

Data flow: request -> token check -> fast path only (create/detect job folder per the fresh-vs-existing logic, save attachment bytes, append/update registry entry with `source="outlook"`) -> returns immediately. No RADAN work is triggered from this listener — that stays a desktop-app action the user drives from the Job Intake tab, avoiding blocking the listener's thread on slow RADAN COM automation.

## Phase 3 — Outlook add-in

**New files** (served as static/template routes by the Phase 2 listener):
- `master_app/static/job_intake_addin/manifest.xml` — sideloaded add-in manifest, ribbon button on a received email opening a task pane.
- `master_app/templates/job_intake_taskpane.html` + `static/job_intake_addin/taskpane.js` — `Office.onReady`, a job-number input (+ Label input, shown conditionally if the server reports the job folder already exists), reads `Office.context.mailbox.item.attachments` + `getAttachmentContentAsync` (base64) filtered to `.dxf`, POSTs to `/api/job-intake` on the embedded listener with the bearer token injected server-side into the rendered HTML (via Jinja, so the token is never hardcoded client-side or typed by hand).

**Real risk, spike early before building the rest of the task pane:** Office's task-pane webview can be stricter about self-signed certificate trust than a regular browser tab. Confirm the locally-generated root CA (once installed into the Windows trust store on the user's PC) is actually accepted by Office before investing further in Phase 3.

## HANDOFF — PHASE 1 COMPLETE (two master_app commits). Remaining: Phases 2 & 3.

**Phase 1 is fully built, committed, and verified.** Commit `113ea63` = service layer; the follow-up commit = the Job Intake tab UI + wiring + tests. `ops_job_intake.py` (`JobIntakePage`), wired into `ops_main_window.py` as sidebar index 2 (Diagnostics moved to index 3; `_select_page` and the nav `labels` tuple updated). 65 tests pass. `PyMuPDF` added to `requirements.txt`. Live-verified by driving the real `OpsMainWindow` offscreen: tab renders, Manual-Intake -> Create-RPD produces `M59919.rpd` in the `M59919/M59919/nests` shape. `_create_intake(job_number, label, files)` on the page is the testable seam the dialog calls; the Phase 2 listener will call the same `job_intake_service`/`job_intake_registry` functions with `source="outlook"`. Registry funcs resolve their path at call time (monkeypatch-friendly). PO extraction also returns `due_note` for RUSH/ASAP.

**Remaining: Phase 2 (embedded listener) and Phase 3 (Outlook add-in)** — see the sections further below; nothing in the service layer needs to change for them.

### Earlier handoff detail (service layer, still accurate):
**Done, committed, and validated (master_app commit `113ea63`):**
- `ops_paths.py`: `BATTLESHIELD_ROOT`, `MACHINE_EIA_BATTLESHIELD_ROOT`, `JOB_PREFIX_TO_ROOT`, `JOB_INTAKE_REGISTRY_PATH`.
- `job_intake_registry.py`: full registry (new_entry/append/get/update/delete/load_entries newest-first, `entry_key(job_number, label)` composite keys, status constants + validation, atomic writes).
- `job_intake_service.py`: `resolve_job_root`/`resolve_job_paths`/`job_folder_exists`/`create_job_folders` (fresh-vs-Label branches enforced — fresh raises if the folder exists, directing the user to add a Label), `clone_rpd_template` (standalone, takes `template_path` param defaulting to `EXPLORER_TEMPLATE_PATH`), `material_choices()`/`default_strategy_for_material()`, `copy_attachments`, `extract_po_hints`, `build_import_csv_rows`/`write_import_csv`, `launch_radan_import(explorer_services, ...)`, `send_job_blocks_to_machine(explorer_services, ...)`.
- `tests/test_job_intake.py`: 17 tests, all passing (61/61 suite-wide). PO tests build synthetic PDFs with PyMuPDF mimicking the verified cell-per-line layout.
- **PO extraction was validated against five real POs on L: spanning 2022–2026** (M59919, M59841, M50087, M50239, M59761) — PO number, due date, per-DXF qty, and unmatched-line reporting all correct, including the hard cases (L1 vs L1-CP, End Cap D vs D_2, "Channel - 11a" vs `Channel 11a.dxf`, RUSH -> no due date, PO-/Q- prefixes, -R1 suffixes). Critical implementation fact: **PyMuPDF linearizes the PO table one cell per line** (`"1"`, `"10"`, `"Clip-End - 1/4\" Mild Steel"`), NOT row-per-line — the matcher is description-anchored and reads qty from the two integer lines above; do not "simplify" it back to a row regex.
- Known accepted limitation: a short stem (`M1`) prefix-claims rows `M10`–`M19` when those DXFs aren't attached, so `unmatched_lines` can undercount; the warning still fires on the remaining lines.

**Verified integration facts (do NOT re-explore):**
- `truck_nest_explorer/services.py` **re-exports** `launch_radan_csv_import` and `send_project_block_files_to_machine`, and `services` is in `TRUCK_EXPLORER_MODULES` — so `load_truck_explorer_api().services.<fn>` reaches both. **No `source_apps.py` changes needed.**
- master_app tests are pytest with a module-scoped `qapp` fixture and `QT_QPA_PLATFORM=offscreen` (see `tests/test_ops_main_window_navigation.py`). That test builds `OpsMainWindow` with a `SimpleNamespace` fake explorer_api — therefore **`JobIntakePage` must not touch `explorer_api` during construction** (only inside button handlers), or the nav test breaks.
- Shared venv (`C:\Tools\.venv`) already has PyMuPDF (`fitz`) — but add `PyMuPDF` to `master_app/requirements.txt` for the record.
- `git` in master_app: commit (auto-commit is approved per user memory), do not push.

**Remaining Phase 1 work (mechanical, patterns all established):**
1. `master_app/ops_job_intake.py` — `JobIntakePage(QWidget)` modeled on `ops_admin.py`'s `DiagnosticsPage` (same header/title/refresh shape, `page_title`/`page_subtitle`/`panel` object names for QSS). Layout: queue table (left or top; columns Job #, Label, PO #, Status, Due, Received, Source) + detail panel: job fields (Due Date via QDateEdit with a "no date" checkbox or blank sentinel, Laser hours + Bend hours QDoubleSpinBox 0–999 × 0.25 step, PO number shown as label), an unmatched-PO-lines warning banner (visible when the selected entry has stored unmatched lines — persist them on the entry dict as e.g. `po_unmatched: [...]` when creating the intake), and the parts grid (QTableWidget one row per DXF attachment): Material (editable combo from `material_choices()`, delegate modeled on `radan_kitter/ui_parts_table.py:KitComboDelegate`), Thickness (QDoubleSpinBox delegate, 3 decimals, min 0.001), Qty (QSpinBox delegate 1–9999), Unit (combo from `UNIT_CHOICES`), Strategy (read-only, auto-refreshed from material via `default_strategy_for_material`), PO Ref (read-only `raw_description`). Buttons: **Manual Intake** (dialog: job number field -> on submit call `job_folder_exists`; if True show/require Label field; QFileDialog multi-select for DXF+PDF; then `resolve_job_paths` -> `create_job_folders` -> `copy_attachments` -> `extract_po_hints` on every non-DXF PDF (merge: first non-None po_number/due_date wins; union line_items; union unmatched) -> seed `material_qty` rows (one per DXF, qty from hints else 1, unit "in", material blank) -> `new_entry` + `append_entry` with `job_folder`/`po_number`/`due_date` set), **Save Details** (`update_entry` with grid + job fields), **Create Blank RPD** (`clone_rpd_template`, status -> `rpd_created`, store `rpd_path`), **Import Parts to RADAN** (auto-save, `build_import_csv_rows` -> `write_import_csv` to `<intake_dir>/<project_name>-BOM_Radan.csv` -> `launch_radan_import` with `log_path=_runtime/job_intake_radan_import_<key>_<ts>.log`; keep the Popen, poll it in the page timer; rc==0 -> status `parts_imported`, else status `error` + error message pointing at the log), **Send Blocks to Machine** (confirm dialog, run `send_job_blocks_to_machine` in a `concurrent.futures.ThreadPoolExecutor(1)` future polled by the timer, status -> `blocks_sent`; enabled once status is `parts_imported` or later since nesting happens manually in RADAN between import and blocks), **Refresh**. A ~4s `QTimer` polls the registry + any live Popen/future, running only while the page is visible (start in `showEvent`, stop in `hideEvent`).
2. Wire into `ops_main_window.py`: construct `JobIntakePage(explorer_api=self._explorer_api)` alongside `DiagnosticsPage` (~line 231), add to the page tuple (~237–242), add "Job Intake" to the sidebar `labels` tuple (line 267), and in `_select_page` refresh it on show like index 2 does for diagnostics (the diagnostics hook indexes shift if Job Intake is inserted before it — keep Diagnostics last, Job Intake as index 2, and update the `_select_page` index checks accordingly, plus the nav test's expectations).
3. `requirements.txt`: add `PyMuPDF`.
4. Tests: extend `test_ops_main_window_navigation.py` for the 4th page (fake explorer_api unchanged — page must construct without touching it), plus a `JobIntakePage` test in the offscreen style: build the page with a `SimpleNamespace` services fake, monkeypatch `job_intake_service.BATTLESHIELD_ROOT` and `JOB_INTAKE_REGISTRY_PATH` to tmp, drive Manual Intake logic via the page's non-dialog helper methods (structure the page so the dialog collects inputs and a testable `_create_intake(job_number, label, files)` method does the work), assert the entry appears in the queue and Create RPD produces the file and status transition.
5. Live verification per the Phase 1 steps below; then commit.

**Phases 2–3 are unchanged** (see sections above); Phase 2's listener reuses `job_intake_service`/`job_intake_registry` exactly as built — the `source="outlook"` path is `new_entry(source="outlook", ...)` + the same create/copy/extract sequence, nothing new in the service layer.

## HANDOFF — PHASE 2 COMPLETE. Remaining: Phase 3 (Outlook add-in).

**Built and live-verified over a real HTTPS socket.** Three new files:

- `job_intake_service.create_intake(job_number, label, files, *, source, email_subject, email_sender)` — the intake sequence was **moved off the Qt page into the service**. It is now the single copy: `JobIntakePage._create_intake` delegates to it with `source="manual"`, the listener calls it with `source="outlook"`. Do not reintroduce the sequence in a caller.
- `job_intake_tls.py` — self-contained local root CA + leaf cert. Self-contained on purpose: master_app has a similar `web_tls`, but master_app now *embeds* this repo, so importing it back would be a dependency cycle.
- `job_intake_server.py` — Flask app on `127.0.0.1` only, bearer token from `_runtime/job_intake_api_token.key`, started as a daemon thread from `app.py` (disable with `ODD_JOB_INTAKE_LISTENER=0`, port via `ODD_JOB_INTAKE_PORT`, default 8790).

Routes: `POST /api/job-intake` (job number + base64 attachments → 201 with the entry summary), `GET /api/job-intake/check?job_number=` (→ `label_required`, so the task pane knows whether to show its Label field), `GET /api/health` (**unauthenticated by design** — lets the task pane distinguish "desktop app isn't running" from "bad token"), `GET /job-intake-root-ca.crt` (for the Windows trust-store install).

**Settled implementation facts:**
- **The listener triggers no RADAN work**, as designed — it creates the folder, decodes attachments, scrapes the PO, and registers the entry, then returns. A test asserts no `.rpd` is produced. RPD/import/blocks stay desktop actions.
- Attachments are accepted only as `.dxf`/`.pdf`, and each name is reduced to a bare filename before being joined to a path under `L:` — an emailed attachment name is untrusted input and `..\..\` traversal is stripped (tested).
- At least one `.dxf` is required; a PDF-only submission is refused.
- **Resubmitting a job returns 400 asking for a Label, not 409** — `create_job_folders`'s existing-folder guard fires before the registry's duplicate-key check. The 409 branch is only reachable when the folder was deleted by hand but the registry entry remains.
- Uses Flask's threaded WSGI server, **not waitress** (which has no built-in TLS termination). Correct for one user's occasional single requests on loopback.
- `ensure_loopback_certificate()` is idempotent and reuses existing material — it must be, or every startup would invalidate the CA the user installed into Windows.

### Phase 3 cert spike — NOT a blocker, but there are two extra steps

The flagged risk was that Office's webview might reject a self-signed cert. It does not, **provided the root CA is installed into the Windows "Trusted Root Certification Authorities" store** (per-user is enough). Two findings that are easy to lose a day to:

1. **The cert must cover both `localhost` (DNS) and `127.0.0.1` (IP).** Office accepts either spelling in a manifest's `SourceLocation`, and a cert issued for only one fails on the other — this is a known footgun (OfficeDev/Office-Addin-Scripts#514). `job_intake_tls` puts both in the SAN and a test pins it.
2. **WebView2 needs a loopback exemption.** The Office webview app container blocks loopback by default; without this the task pane fails with an opaque "we can't open this add-in from localhost". In an **admin** PowerShell:
   ```
   CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
   ```
   Checked on this machine (2026-07-22): **no exemption is currently registered**, so this step is still outstanding. It is a system change requiring admin, so it was documented rather than applied.

Sanity check before writing any add-in code: open `https://127.0.0.1:8790/api/health` in Edge — no cert warning means the trust half is done.

Because the task pane HTML will be served by this same listener, the pane is same-origin with the API and **no CORS configuration is needed**. Serving it from elsewhere would change that.

## Future work — better material/qty extraction when there is no PO

Today extraction only runs against non-DXF PDF attachments, so a job emailed without a PO document falls back to fully manual entry. Worth pursuing, roughly in order of expected payoff:

- **Scrape the email's own plain-text body.** Often the material and quantities are simply typed in the message ("2 off, 1/4 mild steel") when no formal PO is attached. The listener already receives the email context, so the body would just become another `POHints` source. Material stays reference-only per the existing data-entry rule — the same inconsistent-spelling problem applies, and arguably worse in free-form prose.
- **PDF drawing prints, not just POs.** Job folders routinely contain a second, non-PO PDF (the drawings), and its title block carries material/thickness per part. `truck_nest_explorer` already does title-block text extraction with PyMuPDF for a similar purpose — worth looking at before writing anything new.
- **Text embedded in the DXF/DWG itself.** TEXT/MTEXT entities and the drawing's own title block frequently name material and thickness. This is the most reliable source of the three when present (it is the designer's own value, not a salesperson's retyping) and the only one available when the DXF arrives with no accompanying document at all. DXF is plain text and parseable directly; DWG would need a converter.

All three feed the same best-effort contract: pre-fill Qty (and possibly show thickness as reference), never auto-fill Material, and surface anything unmatched rather than guessing.

## Verification

- **Phase 1**: run the desktop app (`dev_run.bat`), open the new Job Intake tab, use "Manual Intake" with a real small DXF set and a test job number (e.g. a fresh `M9####...` number and, separately, an existing job number to exercise the Label-required branch), fill in material/thickness/qty/due-date/estimates, generate the RPD, confirm the resulting `.rpd`/folder shape matches the `M59919` convention exactly, confirm parts appear correctly in RADAN, and test "Send Blocks to Machine" once nests/`.cnc` files exist.
- **Phase 2** (done): verified against a real HTTPS socket on 127.0.0.1 with TLS validated against the generated root CA — health, token rejection, a two-DXF submission landing in the correct `M-FABRICATION/<job>` shape with `source="outlook"` and no `.rpd`, the resubmit guard, and the root-CA download route. The `localhost` spelling was confirmed to validate against the same cert. Still worth doing by hand once against the real `L:` drive with the desktop app running, to confirm the intake appears in the tab's queue on the next poll.
- **Phase 3**: sideload the manifest in Outlook (desktop or web), confirm the task pane loads over HTTPS without a cert warning (post-trust), submit a real test email's DXF attachments, and confirm the same end-to-end flow as Phase 1/2 completes.
- Run each existing test suite (`master_app/tests/`, `truck_nest_explorer/tests/`) after each phase to confirm no regressions in the embedding/navigation code touched.
