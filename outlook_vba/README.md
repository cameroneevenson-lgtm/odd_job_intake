# Job Intake button for classic Outlook (VBA)

A ribbon button that sends the selected email's DXF/PDF attachments to the
loopback listener. Use this **only because** Outlook web add-ins install
through Exchange and this tenant blocks that — see `../docs/SIDELOAD.md`.

`JobIntake.bas` is transport only. It holds no intake logic: it calls the same
`/api/job-intake` the web task pane would, which runs the same
`job_intake_service.create_intake` the desktop app uses. Keep it that way.

## Why VBA and not a COM/VSTO add-in

Checked on this machine: no Visual Studio, no MSBuild, no VSTO runtime. A COM
add-in would mean installing a VSTO runtime (admin), adding a build toolchain,
and maintaining a C# project in an otherwise all-Python toolshed — **for the
same shelf life**, since new Outlook drops COM and VBA alike. The durable path
is the web add-in, which is already built and waiting on a tenant flag.

## Install

### 1. Import the module

Outlook → **Alt+F11** → File → **Import File…** → pick `JobIntake.bas` →
**Ctrl+S**.

This creates `%APPDATA%\Microsoft\Outlook\VbaProject.OTM` (there is no macro
project on this machine yet). Note that file is **not** version controlled —
re-import from here after an Outlook profile rebuild.

### 2. Deal with macro security — do this, or the button silently won't run

Outlook's default is *"Notifications for digitally signed macros, all other
macros disabled"*, so an unsigned macro is **blocked with no error**. Pick one:

**Preferred — sign it** (keeps security at the default):

1. Run `C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE`
2. Name the certificate e.g. `Battleshield Macros` → OK
3. Back in the VBA editor: Tools → **Digital Signature…** → Choose → pick that
   certificate → OK → **Ctrl+S**
4. Restart Outlook. On first run choose **Trust all documents from this
   publisher** if prompted.

**Alternative — lower the setting** (weakens it for *all* macros, not just
this one): File → Options → Trust Center → Trust Center Settings → Macro
Settings → *Notifications for all macros*.

### 3. Put it on the ribbon

Right-click the ribbon → **Customize the Ribbon** → in the right pane select
the **Home (Mail)** tab → **New Group** → set *Choose commands from* to
**Macros** → add `Project1.JobIntake.SendToJobIntake` → **Rename…** it to
`Send to Job Intake` and pick an icon → OK.

(Quick Access Toolbar works too and is fewer clicks, if you don't mind it
living top-left.)

## Use

1. The **Ops Suite must be running** — the listener lives inside it and is
   loopback-only. The macro checks this first and says so plainly if not.
2. Select a received email with `.dxf` attachments.
3. Click **Send to Job Intake**.
4. Enter the job number. If that job already has a folder on `L:`, it asks for
   a Label so the one-off gets its own subfolder — the same rule the desktop
   app applies.
5. A summary shows the PO number, due date, part count, and folder.

**No RADAN work runs.** The email only stages files and registers the intake.
Material, thickness, RPD creation, part import, and block transfer stay in the
Job Intake tab — deliberately, so a mail action can never block on RADAN COM
automation.

## Troubleshooting

| Symptom | Cause |
| --- | --- |
| Nothing happens at all | Macro blocked by security — do step 2 |
| "The shop app isn't running on this PC" | Ops Suite closed; the listener dies with it |
| "Could not read the API token" | Start the Ops Suite once to generate `_runtime/job_intake_api_token.key` |
| Certificate/connection error | Root CA missing from the trust store — see `../docs/SIDELOAD.md` |
| "no .dxf attachment" | Correct: a PDF-only email has nothing to nest |

The macro reads the port from `ODD_JOB_INTAKE_PORT` and falls back to 8790,
matching the listener's own default — don't hardcode a second value.

## Security notes

- The API token is read at run time from `_runtime/`, never stored in the
  macro. That directory is gitignored; this repo is **public**.
- Attachments are written to `%TEMP%\job_intake_vba\` only long enough to
  base64-encode, then deleted.
- Attachment names are sanitised before being used as paths, because they come
  from email. The server does this again independently.
- The macro never disables certificate validation. A TLS failure means the
  root CA is missing and should be fixed, not bypassed.
