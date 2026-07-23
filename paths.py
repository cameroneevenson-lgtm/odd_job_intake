"""Shop paths and constants for odd_job_intake.

Standalone version of the handful of constants this feature used while it
lived inside master_app's ops_paths.py. Kept in one place so both the desktop
page and the (planned) Outlook-facing listener resolve identical locations.
"""

from __future__ import annotations

import os
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
TOOLS_DIR = Path(os.environ.get("ODD_JOB_INTAKE_TOOLS_DIR", r"C:\Tools"))

# Sibling apps this feature borrows behavior from at runtime.
TRUCK_NEST_EXPLORER_DIR = TOOLS_DIR / "truck_nest_explorer"
INVENTOR_TO_RADAN_DIR = TOOLS_DIR / "inventor_to_radan"

# The blank RADAN project cloned for each new job.
EXPLORER_TEMPLATE_PATH = TRUCK_NEST_EXPLORER_DIR / "Template" / "Template.rpd"

# One-off jobs live under the shop's existing L: roots, chosen by the job
# number's prefix letter - never a new root folder.
BATTLESHIELD_ROOT = Path(r"L:\BATTLESHIELD")
MACHINE_EIA_BATTLESHIELD_ROOT = Path(r"A:\EiaFiles\Battleshield")
JOB_PREFIX_TO_ROOT = {
    "F": "F-LARGE FLEET",
    "P": "P-SMALL FLEET",
    "M": "M-FABRICATION",
    "W": "W-WARRANTY",
    "S": "S-SERVICE",
}

JOB_INTAKE_REGISTRY_PATH = APP_DIR / "_runtime" / "job_intake_registry.json"

# Where an email is allowed to point. A message body is untrusted text, and a
# path lifted out of one is followed and read, so it is restricted to the two
# shares this work legitimately comes from: engineering's laser folder on W:
# and the shop's own job roots on L:. Anything else is ignored rather than
# refused, because prose contains paths that were never an instruction.
APPROVED_SOURCE_ROOTS = (
    Path(r"W:\LASER"),
    BATTLESHIELD_ROOT,
)

# Placeholder numbers for work that arrives before it has been assigned one -
# "start looking at it, the number comes later". The prefix still picks the
# root, so an M placeholder lands under M-FABRICATION like any other.
#
# One per prefix, and every intake filed against one is *required* to carry a
# label, so each parks in its own subfolder (M12345\PFF PO-8527-001\) instead
# of several unrelated jobs piling into the same directory. Rename Job moves
# it to the real number once that exists.
PLACEHOLDER_JOB_NUMBERS = frozenset(
    {f"{prefix}12345" for prefix in JOB_PREFIX_TO_ROOT}
)
