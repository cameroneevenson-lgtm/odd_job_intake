from __future__ import annotations

import csv
import dataclasses
from datetime import date
import json
from pathlib import Path
import threading

import pytest

import job_intake_registry
import job_intake_service
from job_intake_registry import (
    STATUS_RPD_CREATED,
    append_entry,
    delete_entry,
    entry_key,
    get_entry,
    load_entries,
    new_entry,
    update_entry,
)
from job_intake_service import (
    JobIntakeError,
    build_import_csv_rows,
    clone_rpd_template,
    create_job_folders,
    default_strategy_for_material,
    extract_po_hints,
    material_choices,
    resolve_job_paths,
    resolve_job_root,
    write_import_csv,
)


# --- registry ----------------------------------------------------------------


def test_registry_append_get_update_delete_round_trip(tmp_path: Path) -> None:
    registry_path = tmp_path / "registry.json"
    entry = new_entry(job_number="m59919", source="manual")
    append_entry(entry, registry_path)

    loaded = get_entry("M59919", registry_path)
    assert loaded is not None
    assert loaded["job_number"] == "M59919"
    assert loaded["status"] == "new"

    update_entry("M59919", registry_path, status=STATUS_RPD_CREATED, rpd_path="x.rpd")
    updated = get_entry("M59919", registry_path)
    assert updated is not None
    assert updated["status"] == STATUS_RPD_CREATED
    assert updated["rpd_path"] == "x.rpd"

    delete_entry("M59919", registry_path)
    assert get_entry("M59919", registry_path) is None


def test_registry_rejects_duplicate_keys_and_bad_status(tmp_path: Path) -> None:
    registry_path = tmp_path / "registry.json"
    append_entry(new_entry(job_number="M59919"), registry_path)
    with pytest.raises(ValueError):
        append_entry(new_entry(job_number="M59919"), registry_path)
    with pytest.raises(ValueError):
        update_entry("M59919", registry_path, status="not-a-status")
    # A labeled one-off under the same number is a distinct entry.
    append_entry(new_entry(job_number="M59919", label="Rush Plates"), registry_path)
    assert entry_key("M59919", "Rush Plates") == "M59919::rush plates"
    assert len(load_entries(registry_path)) == 2


def test_registry_load_entries_newest_first(tmp_path: Path) -> None:
    registry_path = tmp_path / "registry.json"
    first = new_entry(job_number="M50001")
    first["received_at"] = "2026-01-01T08:00:00"
    second = new_entry(job_number="M50002")
    second["received_at"] = "2026-07-01T08:00:00"
    append_entry(first, registry_path)
    append_entry(second, registry_path)
    assert [entry["job_number"] for entry in load_entries(registry_path)] == ["M50002", "M50001"]


def test_concurrent_updates_to_different_fields_do_not_lose_each_other(
    tmp_path: Path,
) -> None:
    """The race the JSON store lost.

    Three writers touch this registry: the Flask request thread, its background
    extraction worker, and the Qt UI. Read-modify-write over a whole JSON file
    meant two of them saving at once ended with the second overwriting the
    first's field. Each update now runs inside its own transaction, so every
    write survives regardless of interleaving.
    """
    registry_path = tmp_path / "registry.json"
    append_entry(new_entry(job_number="M51000"), registry_path)

    errors: list[BaseException] = []
    start = threading.Barrier(8)

    def write(index: int) -> None:
        try:
            start.wait(timeout=10)
            update_entry("M51000", registry_path, **{f"field_{index}": index})
        except BaseException as exc:  # surfaced on the main thread below
            errors.append(exc)

    threads = [threading.Thread(target=write, args=(index,)) for index in range(8)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=30)

    assert not errors, errors
    entry = get_entry("M51000", registry_path)
    assert entry is not None
    for index in range(8):
        assert entry[f"field_{index}"] == index, f"update {index} was lost"


def test_two_intakes_racing_for_one_job_number_leave_exactly_one_winner(
    tmp_path: Path,
) -> None:
    """Claiming a job is what stops the same one being filed twice.

    The old check was "load, look for the key, append", which two callers could
    both pass before either wrote. The primary key decides it instead: one
    INSERT wins and the rest get the same ValueError the check used to raise.
    """
    registry_path = tmp_path / "registry.json"
    claimed: list[str] = []
    refused: list[str] = []
    start = threading.Barrier(6)

    def claim() -> None:
        start.wait(timeout=10)
        try:
            append_entry(new_entry(job_number="M51001"), registry_path)
            claimed.append("won")
        except ValueError:
            refused.append("lost")

    threads = [threading.Thread(target=claim) for _ in range(6)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=30)

    assert len(claimed) == 1, "a job number was claimed more than once"
    assert len(refused) == 5
    assert len(load_entries(registry_path)) == 1


def test_the_old_json_registry_is_carried_across_on_first_use(tmp_path: Path) -> None:
    """Upgrading must not lose jobs already in the queue."""
    registry_path = tmp_path / "registry.json"
    registry_path.write_text(
        json.dumps(
            {
                "version": 1,
                "entries": [
                    {
                        "key": "F59487",
                        "job_number": "F59487",
                        "received_at": "2026-07-01T08:00:00",
                        "status": STATUS_RPD_CREATED,
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    entry = get_entry("F59487", registry_path)
    assert entry is not None
    assert entry["status"] == STATUS_RPD_CREATED

    # Renamed rather than deleted: an upgrade should not destroy the old file,
    # and leaving it in place would re-import entries that have since changed.
    assert not registry_path.exists()
    assert registry_path.with_suffix(".json.migrated").exists()

    update_entry("F59487", registry_path, status=job_intake_registry.STATUS_BLOCKS_SENT)
    assert get_entry("F59487", registry_path)["status"] == job_intake_registry.STATUS_BLOCKS_SENT


def test_an_entry_written_before_states_existed_reads_as_finished(tmp_path: Path) -> None:
    """Those entries were saved synchronously, so they really are done - and
    reporting them as still running would hang a poller on them forever."""
    assert job_intake_registry.entry_state({"key": "M51002"}) == (
        job_intake_registry.STATE_SUCCEEDED
    )
    # A manual intake is likewise complete the moment it is written.
    assert job_intake_registry.entry_state(new_entry(job_number="M51003")) == (
        job_intake_registry.STATE_SUCCEEDED
    )
    assert job_intake_registry.entry_state(
        new_entry(job_number="M51004", source="outlook")
    ) == job_intake_registry.STATE_QUEUED


# --- path resolution ---------------------------------------------------------


def test_resolve_job_root_maps_prefixes_and_rejects_bad_numbers() -> None:
    assert resolve_job_root("M59919") == "M-FABRICATION"
    assert resolve_job_root("w50123") == "W-WARRANTY"
    assert resolve_job_root("S123456") == "S-SERVICE"
    with pytest.raises(JobIntakeError):
        resolve_job_root("X59919")
    with pytest.raises(JobIntakeError):
        resolve_job_root("M59A19")
    with pytest.raises(JobIntakeError):
        resolve_job_root("")


def test_resolve_job_paths_fresh_job_matches_shop_convention() -> None:
    paths = resolve_job_paths("m59919")
    assert paths.intake_dir == paths.job_dir
    assert paths.job_dir.name == "M59919"
    assert paths.job_dir.parent.name == "M-FABRICATION"
    assert paths.project_dir == paths.job_dir / "M59919"
    assert paths.rpd_path == paths.project_dir / "M59919.rpd"


def test_resolve_job_paths_labeled_job_nests_under_label() -> None:
    paths = resolve_job_paths("F55334", "Rush Plates")
    assert paths.intake_dir == paths.job_dir / "Rush Plates"
    assert paths.project_name == "F55334 Rush Plates"
    assert paths.project_dir == paths.intake_dir / "F55334 Rush Plates"
    assert paths.rpd_path.name == "F55334 Rush Plates.rpd"


def test_create_job_folders_requires_label_when_job_exists(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path)
    fresh = resolve_job_paths("M59919")
    create_job_folders(fresh)
    assert (fresh.project_dir / "nests").is_dir()
    assert (fresh.project_dir / "remnants").is_dir()

    # Same number again without a label must refuse instead of mixing in.
    with pytest.raises(JobIntakeError):
        create_job_folders(resolve_job_paths("M59919"))

    labeled = resolve_job_paths("M59919", "Extra Brackets")
    create_job_folders(labeled)
    assert labeled.intake_dir.is_dir()
    with pytest.raises(JobIntakeError):
        create_job_folders(resolve_job_paths("M59919", "Extra Brackets"))


# --- RPD template clone ------------------------------------------------------


TEMPLATE_TEXT = """<?xml version="1.0" encoding="UTF-8"?>
<RadanProject xmlns="http://www.radan.com/ns/project">
  <JobName>Template</JobName>
  <NestFolder>C:\\old\\nests</NestFolder>
  <RemnantSaveFolder>C:\\old\\remnants</RemnantSaveFolder>
  <Part><Symbol>Template.rpd</Symbol></Part>
</RadanProject>
"""


def test_clone_rpd_template_substitutes_job_values(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path)
    template_path = tmp_path / "Template.rpd"
    template_path.write_text(TEMPLATE_TEXT, encoding="utf-8")

    paths = resolve_job_paths("M59919")
    create_job_folders(paths)
    rpd_path = clone_rpd_template(paths, template_path)

    text = rpd_path.read_text(encoding="utf-8")
    assert "<JobName>M59919</JobName>" in text
    assert f"<NestFolder>{paths.project_dir / 'nests'}</NestFolder>" in text
    assert f"<RemnantSaveFolder>{paths.project_dir / 'remnants'}</RemnantSaveFolder>" in text
    assert "<Symbol>M59919.rpd</Symbol>" in text
    assert "Template" not in text

    with pytest.raises(JobIntakeError):
        clone_rpd_template(paths, template_path)


# --- material list -----------------------------------------------------------


def test_material_choices_reads_rules_and_excludes_ftq(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", tmp_path)
    (tmp_path / "description_rules.csv").write_text(
        "Description,Material,Thickness,Strategy\n"
        "A,Aluminum 5052,0.12,Air\n"
        "B,Aluminum 3003 CHK FTQ,0.12,Air\n"
        "C,Mild Steel-A36,0.25,O2\n"
        "D,aluminum 5052,0.18,Air\n",
        encoding="utf-8",
    )
    choices = material_choices()
    assert "Aluminum 5052" in choices
    assert "Mild Steel-A36" in choices
    assert not any("FTQ" in choice for choice in choices)
    assert len([c for c in choices if c.casefold() == "aluminum 5052"]) == 1


def _write_shop_csvs(directory: Path, descriptions: list[str], rules: list[tuple[str, str, str, str]]) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    with (directory / "expected_laser_descriptions.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Description"])
        for description in descriptions:
            writer.writerow([description])
    with (directory / "description_rules.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Description", "Material", "Thickness", "Strategy"])
        for row in rules:
            writer.writerow(list(row))


def test_catalog_offers_only_thicknesses_valid_for_each_material(tmp_path: Path, monkeypatch) -> None:
    """The catalog comes from description_rules.csv, and a material/thickness
    pair it doesn't list must not be selectable.

    expected_laser_descriptions.csv is deliberately not consulted - it lists a
    far narrower set, and gating on it rejected 3003 checker plate that came
    through a real customer BOM.
    """
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=[
            'SHEET, AL ALY, .125" THK, 5052 H32',
            'PLATE, AL ALY, .375" THK, 5052 H32',
            'PLATE, MS, .375" THK, 44W',
            'SHEET, AL ALY, .125" THK, 3003 H22 APT FTQ',
        ],
        rules=[
            ('SHEET, AL ALY, .125" THK, 5052 H32', "Aluminum 5052", "0.12", "Air"),
            ('PLATE, AL ALY, .375" THK, 5052 H32', "Aluminum 5052", "0.38", "Air"),
            ('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2"),
            ('SHEET, AL ALY, .125" THK, 3003 H22 APT FTQ', "Aluminum 3003 CHK FTQ", "0.18", "Air"),
            # In the rules but not in the expected list - must still be offered,
            # since the expected list is no longer a gate.
            ("PLATE, SS, .25 THK, 304", "Stainless Steel", "0.25", "N2"),
        ],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)

    assert job_intake_service.material_choices() == (
        "Aluminum 5052",
        "Mild Steel-A36",
        "Stainless Steel",
    )
    # FTQ is a forced per-part override elsewhere, never a user choice.
    assert not any("FTQ" in material for material in job_intake_service.material_choices())

    assert job_intake_service.thickness_choices("Aluminum 5052") == (0.12, 0.38)
    assert job_intake_service.thickness_choices("Mild Steel-A36") == (0.375,)
    # 0.375 is valid for steel but must not be offered for aluminium.
    assert 0.375 not in job_intake_service.thickness_choices("Aluminum 5052")
    assert job_intake_service.thickness_choices("nothing like this") == ()


def test_catalog_picks_up_materials_added_after_launch(tmp_path: Path, monkeypatch) -> None:
    """The shop edits this CSV; new materials must appear without a restart,
    so nothing may be cached at import time."""
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=['PLATE, MS, .375" THK, 44W'],
        rules=[('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2")],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)
    assert job_intake_service.material_choices() == ("Mild Steel-A36",)

    # The shop adds a material mid-session.
    _write_shop_csvs(
        shop,
        descriptions=['PLATE, MS, .375" THK, 44W', "PLATE, SS, .25 THK, 304"],
        rules=[
            ('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2"),
            ("PLATE, SS, .25 THK, 304", "Stainless Steel", "0.25", "N2"),
        ],
    )
    assert job_intake_service.material_choices() == ("Mild Steel-A36", "Stainless Steel")
    assert job_intake_service.thickness_choices("Stainless Steel") == (0.25,)


@pytest.fixture()
def material_memory(tmp_path: Path, monkeypatch) -> Path:
    """Isolate the learned-wording store; it must never touch real _runtime."""
    path = tmp_path / "material_fingerprints.json"
    monkeypatch.setattr(job_intake_service, "MATERIAL_MEMORY_PATH", path)
    return path


# --- material fingerprinting -------------------------------------------------
# A hash that is *supposed* to collide: every way of writing one material has
# to land in the same bucket, or the learned mapping is useless.


@pytest.mark.parametrize(
    "variants",
    [
        pytest.param(
            [
                "MATL: ALUMINIUM 5052",
                "AL ALY 5052-H32",
                "5052 alum plate",
                "aluminum 5052",
                '1/4 THK AL 5052',
                'SHEET, AL ALY, .125" THK, 5052 H32',
            ],
            id="aluminium-5052",
        ),
        pytest.param(
            ["MILD STEEL", "MS", "M.S.", "CRS", "cold rolled steel", "CARBON STEEL", "hot rolled"],
            id="mild-steel",
        ),
        pytest.param(["A36", "A-36", "ASTM A36 STEEL", "MATL: A 36"], id="a36"),
        pytest.param(["44W", "44 W", "CSA 44W PLATE", "MATERIAL: 44W"], id="44w"),
        pytest.param(["SS 304", "304 STAINLESS STEEL", "stainless 304"], id="stainless-304"),
    ],
)
def test_every_spelling_of_one_material_collides(variants) -> None:
    fingerprints = {job_intake_service.material_fingerprint(text) for text in variants}
    assert len(fingerprints) == 1, f"should have collided, got {fingerprints}"
    assert fingerprints != {""}


@pytest.mark.parametrize(
    "text",
    ["SEE NOTE 3", "PART IS 80 X 120", "QTY: 4", '1/4" THK', "GUSSET PLATE", "80 X 120 X 6", ""],
    ids=["note", "dimensions", "qty", "thickness", "part-name", "three-dims", "empty"],
)
def test_text_with_no_material_fingerprints_to_nothing(text) -> None:
    """An empty fingerprint means "said nothing useful", which must not be
    confused with a real bucket - otherwise every unrelated note would collide
    into one and be learned as a material."""
    assert job_intake_service.material_fingerprint(text) == ""


def test_different_materials_do_not_collide() -> None:
    distinct = [
        "ALUMINUM 5052",
        "MILD STEEL",
        "STAINLESS 304",
        "3003 CHECKER PLATE",
    ]
    fingerprints = [job_intake_service.material_fingerprint(text) for text in distinct]
    assert len(set(fingerprints)) == len(distinct), fingerprints


def test_thickness_and_size_never_reach_the_fingerprint() -> None:
    """The same material at different thicknesses is still the same material."""
    quarter = job_intake_service.material_fingerprint('1/4" THK ALUMINUM 5052')
    half = job_intake_service.material_fingerprint('.500 THK ALUMINUM 5052')
    sized = job_intake_service.material_fingerprint("ALUMINUM 5052 80 X 120")
    assert quarter == half == sized


# --- learning from verifications ---------------------------------------------


def test_a_verified_wording_is_recalled_for_every_colliding_spelling(
    material_memory, shop_csvs
) -> None:
    """The point of making the user verify: their confirmation teaches the
    bucket, so all the other spellings of it are predicted next time."""
    assert job_intake_service.recall_material("MATL: ALUMINIUM 5052") is None

    job_intake_service.learn_material_fingerprint("MATL: ALUMINIUM 5052", "Aluminum 5052")

    assert job_intake_service.recall_material("MATL: ALUMINIUM 5052") == "Aluminum 5052"
    # A wording never seen before, but which collides onto the same bucket.
    assert job_intake_service.recall_material("5052 alum plate") == "Aluminum 5052"
    assert job_intake_service.recall_material('AL ALY 5052 H32') == "Aluminum 5052"


def test_learning_beats_the_hand_seeded_aliases(material_memory, shop_csvs) -> None:
    """Evidence from this shop's own drawings outranks a guess about wording."""
    # "CRS" is aliased to Mild Steel-A36 out of the box.
    assert job_intake_service._match_material_in_text("MATL: CRS") == "Mild Steel-A36"

    job_intake_service.learn_material_fingerprint("MATL: CRS", "Aluminum 5052")

    assert job_intake_service._match_material_in_text("MATL: CRS") == "Aluminum 5052"


def test_a_conflicting_bucket_predicts_nothing_and_is_reported(
    material_memory, shop_csvs
) -> None:
    """Two different materials verified onto one fingerprint is a collision
    that shouldn't have happened - guessing between them is worse than asking."""
    job_intake_service.learn_material_fingerprint("MATL: MS", "Mild Steel-A36")
    job_intake_service.learn_material_fingerprint("MATL: MS", "Aluminum 5052")

    assert job_intake_service.recall_material("MATL: MS") is None

    conflicts = job_intake_service.material_memory_conflicts()
    assert len(conflicts) == 1
    bucket = next(iter(conflicts.values()))
    assert set(bucket) == {"Mild Steel-A36", "Aluminum 5052"}


def test_learning_ignores_text_with_no_material(material_memory, shop_csvs) -> None:
    assert job_intake_service.learn_material_fingerprint("SEE NOTE 3", "Aluminum 5052") == ""
    assert job_intake_service.recall_material("SEE NOTE 3") is None


def test_recall_drops_a_material_the_catalog_no_longer_lists(
    material_memory, shop_csvs, tmp_path, monkeypatch
) -> None:
    """The expected-descriptions file stays authoritative even over learning."""
    job_intake_service.learn_material_fingerprint("MATL: ALUMINIUM 5052", "Aluminum 5052")
    assert job_intake_service.recall_material("MATL: ALUMINIUM 5052") == "Aluminum 5052"

    # The shop drops aluminium from the expected list.
    _write_shop_csvs(
        tmp_path / "inventor_to_radan",
        descriptions=['PLATE, MS, .375" THK, 44W'],
        rules=[('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2")],
    )
    assert job_intake_service.recall_material("MATL: ALUMINIUM 5052") is None


# --- thickness snapping and gauges -------------------------------------------


def test_drawing_decimals_snap_onto_the_catalog_value(tmp_path, monkeypatch) -> None:
    """A drawing says .125 where the catalog says 0.12; same sheet. Blanket
    rounding can't do this - it would also turn .375 into .38 and miss mild
    steel's actual 0.375 entry - so snap to the nearest stocked value.

    Uses its own catalog mirroring the real shop files, where aluminium is
    recorded rounded (0.12/0.18/0.38) and mild steel exact (0.375).
    """
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=[
            'SHEET, AL ALY, .125" THK, 5052 H32',
            'PLATE, AL ALY, .188" THK, 5052 H32',
            'PLATE, AL ALY, .375" THK, 5052 H32',
            'PLATE, MS, .375" THK, 44W',
        ],
        rules=[
            ('SHEET, AL ALY, .125" THK, 5052 H32', "Aluminum 5052", "0.12", "Air"),
            ('PLATE, AL ALY, .188" THK, 5052 H32', "Aluminum 5052", "0.18", "Air"),
            ('PLATE, AL ALY, .375" THK, 5052 H32', "Aluminum 5052", "0.38", "Air"),
            ('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2"),
        ],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)

    assert job_intake_service.snap_thickness(0.125, "Aluminum 5052") == 0.12
    assert job_intake_service.snap_thickness(0.1875, "Aluminum 5052") == 0.18

    # The same nominal 3/8 lands on a different number per material, because
    # the shop's own files record it differently - which is precisely why this
    # snaps to the catalog rather than rounding.
    assert job_intake_service.snap_thickness(0.375, "Aluminum 5052") == 0.38
    assert job_intake_service.snap_thickness(0.375, "Mild Steel-A36") == 0.375
    assert job_intake_service.snap_thickness(0.38, "Mild Steel-A36") == 0.375

    # 11ga steel: .118 nominal, CAM wants .12, everyone calls it 1/8 or .125.
    # All four have to converge on the one catalog entry.
    assert job_intake_service.snap_thickness(0.118, "Aluminum 5052") == 0.12
    assert job_intake_service.snap_thickness(0.125, "Aluminum 5052") == 0.12


def test_a_thickness_that_is_not_stocked_snaps_to_nothing(shop_csvs) -> None:
    assert job_intake_service.snap_thickness(1.5, "Aluminum 5052") is None
    assert job_intake_service.snap_thickness(0.0598, "Aluminum 5052") is None
    assert job_intake_service.snap_thickness(0.25, "Mild Steel-A36") is None
    assert job_intake_service.snap_thickness(None, "Aluminum 5052") is None
    assert job_intake_service.snap_thickness(0.12, "not a material") is None


def test_a_gauge_means_different_thicknesses_per_material() -> None:
    """16ga steel is .0598" but 16ga aluminium is .0508" - a gauge number is
    meaningless until the material is known, which is why it is only converted
    after a material has been matched."""
    assert job_intake_service._gauge_to_inches(16, "Mild Steel-A36") == 0.0598
    assert job_intake_service._gauge_to_inches(16, "Aluminum 5052") == 0.0508
    assert job_intake_service._gauge_to_inches(7, "Mild Steel-A36") == 0.1793
    assert job_intake_service._gauge_to_inches(99, "Aluminum 5052") is None


def test_dxf_thickness_snaps_and_gauges_are_read(tmp_path, shop_csvs) -> None:
    quarter = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "A.dxf", "5052 ALUMINUM", '.125" THK')
    )
    assert quarter.thickness == 0.12          # snapped, not 0.125

    gauge = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "B.dxf", "MILD STEEL", "10 GA")
    )
    # 10ga steel is .1345", which this shop doesn't stock - so nothing is
    # claimed rather than snapping to a thickness they'd have to substitute.
    assert gauge.thickness is None
    assert gauge.material == "Mild Steel-A36"


def test_thickness_is_never_claimed_without_a_material(tmp_path, shop_csvs) -> None:
    hints = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "C.dxf", '1/4" THK', "16 GA")
    )
    assert hints.material is None
    assert hints.thickness is None


def test_a_generic_word_cannot_claim_an_alloy_the_shop_lacks(shop_csvs) -> None:
    """Regression: "ALUM" is aliased to the shop's aluminium, so
    "6061-T6 ALUM" was predicted as Aluminum 5052 - the wrong metal."""
    assert job_intake_service._match_material_in_text("MATL: 6061-T6 ALUM") is None
    assert job_intake_service._match_material_in_text("6061 ALUMINIUM") is None
    # A bare generic word is still fine, and the stocked grade still matches.
    assert job_intake_service._match_material_in_text("MATL: ALUM") == "Aluminum 5052"
    assert job_intake_service._match_material_in_text("5052 ALUM") == "Aluminum 5052"


def _dxf_with_text(path: Path, *texts: str) -> Path:
    """A minimal DXF carrying TEXT entities in the real group-code layout:
    each value is preceded by its group code on its own line, and entity text
    is code 1."""
    lines = ["0", "SECTION", "2", "ENTITIES"]
    for text in texts:
        lines += ["0", "TEXT", "8", "NOTES", "1", text]
    lines += ["0", "ENDSEC", "0", "EOF"]
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


@pytest.fixture()
def shop_csvs(tmp_path: Path, monkeypatch) -> Path:
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=[
            'SHEET, AL ALY, .125" THK, 5052 H32',
            'PLATE, AL ALY, .25" THK, 5052 H32',
            'PLATE, MS, .375" THK, 44W',
        ],
        rules=[
            ('SHEET, AL ALY, .125" THK, 5052 H32', "Aluminum 5052", "0.12", "Air"),
            ('PLATE, AL ALY, .25" THK, 5052 H32', "Aluminum 5052", "0.25", "Air"),
            ('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2"),
        ],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)
    return shop


def test_dxf_hints_read_qty_material_and_thickness_from_drawing_text(tmp_path, shop_csvs) -> None:
    path = _dxf_with_text(
        tmp_path / "Gusset.dxf",
        "GUSSET PLATE",
        "MATERIAL: 5052 ALUMINUM",
        'QTY: 4',
        '1/4" THK',
    )
    hints = job_intake_service.extract_dxf_hints(path)

    assert hints.material == "Aluminum 5052"
    assert hints.qty == 4
    # 1/4 -> 0.25, which the shop stocks for this material.
    assert hints.thickness == 0.25
    assert any("QTY" in line for line in hints.raw_lines)


def test_dxf_hints_accept_the_other_qty_wording(tmp_path, shop_csvs) -> None:
    hints = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "Clip.dxf", "2 OFF", "MILD STEEL 44W")
    )
    assert hints.qty == 2
    assert hints.material == "Mild Steel-A36"


def test_dxf_hints_stay_silent_when_the_material_is_ambiguous(tmp_path, shop_csvs) -> None:
    """Material stays the user's choice unless the drawing names exactly one
    material the shop stocks - naming two must not pick a winner."""
    hints = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "Mixed.dxf", "5052 ALUMINUM OR 44W MILD STEEL", "QTY: 3")
    )
    assert hints.material is None
    assert hints.thickness is None      # no material -> no thickness claim
    assert hints.qty == 3               # qty is unambiguous, so it still counts


def test_dxf_hints_reject_a_thickness_the_shop_does_not_stock(tmp_path, shop_csvs) -> None:
    """0.5 is not offered for Aluminum 5052 here, so claiming it would only
    fail RADAN's import later."""
    hints = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "Thick.dxf", "5052 ALUMINUM", '1/2" THK')
    )
    assert hints.material == "Aluminum 5052"
    assert hints.thickness is None


@pytest.mark.parametrize(
    "text",
    ["VALUE", "ASSEMBLY", "ITEMS", "TOTAL VALUE", "SEE ASSEMBLY DRAWING", "ASSY"],
)
def test_ordinary_words_do_not_predict_a_material(text, shop_csvs) -> None:
    """Aliases were matched as bare substrings, so VALUE contained "alu",
    ASSEMBLY contained "ss" and ITEMS contained "ms". Since the email body is
    scraped, that was ordinary prose deciding what gets cut."""
    assert job_intake_service._match_material_in_text(text) is None


@pytest.mark.parametrize(
    "text, expected",
    [
        ("MATL: ALUM", "Aluminum 5052"),
        ("5052 ALUM", "Aluminum 5052"),
        ("cold rolled steel", "Mild Steel-A36"),
        ("MILD STEEL", "Mild Steel-A36"),
        ("MATL: CRS", "Mild Steel-A36"),
        ("44W", "Mild Steel-A36"),
        ("A-36", "Mild Steel-A36"),
        ("304 SS", "Stainless Steel"),
        ("checker plate", "Aluminum 3003 CHK"),
    ],
)
def test_real_material_wording_still_resolves(text, expected) -> None:
    """The word-boundary fix must not cost the multiword aliases - "cold
    rolled" and "checker plate" have to match a run of tokens."""
    assert job_intake_service._match_material_in_text(text) == expected


def test_a_bare_dimension_is_not_mistaken_for_a_material(tmp_path, monkeypatch) -> None:
    """Regression: the real shop file contains `5052 H32 >80"`, whose "80" was
    harvested as an Aluminum token - so any 80mm dimension on a drawing
    claimed aluminium. Numeric tokens must be long enough to be a grade."""
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=['PLATE, AL ALY, .188" THK, 5052 H32 >80"'],
        rules=[('PLATE, AL ALY, .188" THK, 5052 H32 >80"', "Aluminum 5052", "0.18", "Air")],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)

    assert "80" not in job_intake_service._material_tokens()
    assert job_intake_service._match_material_in_text("PART IS 80 X 120") is None
    # The real grade still matches.
    assert job_intake_service._match_material_in_text("MATERIAL 5052") == "Aluminum 5052"


def test_dxf_hints_tolerate_a_drawing_with_nothing_useful(tmp_path, shop_csvs) -> None:
    empty = tmp_path / "Plain.dxf"
    empty.write_text("0\nSECTION\n2\nENTITIES\n0\nENDSEC\n0\nEOF\n", encoding="utf-8")
    hints = job_intake_service.extract_dxf_hints(empty)
    assert (hints.material, hints.thickness, hints.qty) == (None, None, None)

    missing = job_intake_service.extract_dxf_hints(tmp_path / "does_not_exist.dxf")
    assert (missing.material, missing.qty) == (None, None)


def test_dxf_hints_strip_mtext_formatting_codes(tmp_path, shop_csvs) -> None:
    """MTEXT wraps text in formatting markup; the words must survive it."""
    hints = job_intake_service.extract_dxf_hints(
        _dxf_with_text(tmp_path / "Fmt.dxf", r"{\fArial|b1;MATERIAL: 44W} \pxqc;QTY: 6")
    )
    assert hints.material == "Mild Steel-A36"
    assert hints.qty == 6


def test_po_qty_wins_over_the_drawing(tmp_path, monkeypatch, shop_csvs) -> None:
    """The DXF fallback fills gaps; it never overrides what the PO said."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    source = _dxf_with_text(tmp_path / "Bracket.dxf", "QTY: 9", "44W MILD STEEL")

    monkeypatch.setattr(
        job_intake_service,
        "extract_po_hints",
        lambda *_args, **_kwargs: job_intake_service.POHints(
            po_number="8497-005",
            due_date=None,
            due_note=None,
            line_items={"Bracket": {"qty": 25, "raw_description": "Bracket - 3/8 Mild Steel"}},
            unmatched_lines=(),
        ),
    )
    pdf = tmp_path / "po.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")

    entry = job_intake_service.create_intake("M50123", None, [source, pdf])
    part = entry["material_qty"][0]

    assert part["qty"] == 25                       # PO wins
    assert part["material"] == "Mild Steel-A36"    # drawing filled the gap
    assert part["strategy"] == "O2"
    assert "QTY: 9" in part["dxf_ref"]             # what the drawing said, for reference


def test_drawing_qty_used_when_there_is_no_po(tmp_path, monkeypatch, shop_csvs) -> None:
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    source = _dxf_with_text(tmp_path / "Panel.dxf", "QTY: 7", "5052 ALUMINUM", '1/4" THK')

    entry = job_intake_service.create_intake("M50124", None, [source])
    part = entry["material_qty"][0]

    assert part["qty"] == 7
    assert part["material"] == "Aluminum 5052"
    assert part["thickness"] == 0.25
    assert part["strategy"] == "Air"


def test_a_missing_catalog_fails_closed(tmp_path: Path, monkeypatch) -> None:
    """description_rules.csv is the authority for what exists, and it lives on
    C: - unreachable means something is badly wrong. Falling back to a
    hardcoded list would quietly make a stale guess the authority for what gets
    cut, so nothing is offered and the import refuses."""
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", tmp_path / "missing")

    assert material_choices() == ()
    assert job_intake_service.thickness_choices("Aluminum 5052") == ()

    # And the gate refuses rather than letting a previously-filled row through.
    entry = _entry_with_parts(tmp_path)
    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    assert "catalog" in str(excinfo.value).casefold()


def test_import_refuses_a_material_thickness_pair_the_catalog_lacks(
    tmp_path: Path, shop_csvs
) -> None:
    """The last check before anything is written for RADAN, against the catalog
    as it reads now - the file may have changed since the job was filed."""
    entry = _entry_with_parts(tmp_path)
    part = entry["material_qty"][0]
    part["material"] = "Aluminum 5052"
    part["thickness"] = 0.9  # not stocked
    part["resolved"] = {"material": True}

    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    assert "isn't stocked" in str(excinfo.value)

    part["material"] = "Unobtainium"
    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    assert "not in the shop's material list" in str(excinfo.value)


def test_default_strategy_for_material() -> None:
    assert default_strategy_for_material("Aluminum 5052") == "Air"
    assert default_strategy_for_material("Mild Steel-A36") == "O2"
    assert default_strategy_for_material("Stainless Steel") == "N2"
    assert default_strategy_for_material("Stainless Steel 304") == "N2"
    assert default_strategy_for_material("Something New") == "Air"


# --- import CSV --------------------------------------------------------------


def _entry_with_parts(tmp_path: Path) -> dict:
    dxf = tmp_path / "Clip-End.DXF"
    dxf.write_text("dxf", encoding="utf-8")
    entry = new_entry(job_number="M59919")
    entry["attachments"] = [{"filename": "Clip-End.DXF", "saved_path": str(dxf), "size": 3}]
    entry["material_qty"] = [
        {
            "filename": "Clip-End.DXF",
            "material": "Mild Steel-A36",
            "thickness": 0.25,
            "unit": "in",
            "qty": 10,
            "strategy": "",
            # Materials must be human-confirmed before they can reach RADAN.
            "material_confirmed": True,
        }
    ]
    return entry


def _print_pdf(path: Path, lines: list[tuple[float, float, str]]) -> Path:
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    page = doc.new_page()
    for x, y, text in lines:
        page.insert_text((x, y), text)
    doc.save(str(path))
    doc.close()
    return path


def test_a_print_named_by_the_customer_still_matches_its_part(
    tmp_path, monkeypatch, shop_csvs
) -> None:
    """Prints are routinely named and numbered by the customer's own system
    rather than after the part file. The part number still appears somewhere on
    the sheet, so matching only on the title block's drawing-number cell would
    reject the whole drawing."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    dxf = _dxf_with_text(tmp_path / "F57524-C-2.dxf", "GEOMETRY ONLY")
    pdf = _print_pdf(
        tmp_path / "DOC-99871-B.pdf",
        [
            (72, 100, "CUSTOMER DRAWING SET  DOC-99871-B"),
            (72, 130, "ITEM REF: F57524-C-2"),
            (72, 300, "MATERIAL"),
            (72, 310, "Mild Steel 44W"),
        ],
    )

    entry = job_intake_service.create_intake("M90301", None, [dxf, pdf])
    assert entry["material_qty"][0]["material"] == "Mild Steel-A36"


def test_a_po_and_a_print_disagreeing_on_material_is_a_hard_stop(
    tmp_path, monkeypatch, shop_csvs
) -> None:
    """The real case: a PO asking for aluminium against a print drawn in steel.
    The PO's wording never *chooses* the material - customers spell it
    inconsistently - but it absolutely gets a say in whether sources agree."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    dxf = _dxf_with_text(tmp_path / "F57524-C-2.dxf", "GEOMETRY ONLY")
    print_pdf = _print_pdf(
        tmp_path / "DOC-99871-B.pdf",
        [
            (72, 130, "ITEM REF: F57524-C-2"),
            (72, 300, "MATERIAL"),
            (72, 310, "Mild Steel 44W"),
        ],
    )
    po_pdf = _print_pdf(
        tmp_path / "PFF PO-8497-005.pdf",
        [
            (72, 80, "LASER ORDER"),
            (72, 100, "PO Number: 8497-005"),
            (72, 140, "1"),
            (72, 154, "2"),
            (72, 168, 'F57524-C-2 - 1/4" Aluminum 5052'),
        ],
    )

    entry = job_intake_service.create_intake("M90302", None, [dxf, print_pdf, po_pdf])
    part = entry["material_qty"][0]

    conflict = part["conflicts"]["material"]
    assert "the PO says Aluminum 5052" in conflict
    assert "the print says Mild Steel-A36" in conflict

    with pytest.raises(JobIntakeError) as excinfo:
        job_intake_service.build_import_csv_rows(entry)
    assert "STOP" in str(excinfo.value)


def test_files_already_on_the_shared_drive_are_referenced_not_copied(
    tmp_path, monkeypatch, shop_csvs
) -> None:
    """W: is the source of truth for a job filed by path. Copying it to L: only
    created a second version to drift out of date, and was by far the slowest
    part of an intake - a real 28-file job spent 53s on it, past the timeout of
    whatever was waiting."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    shared = tmp_path / "W" / "Job 123"
    shared.mkdir(parents=True)
    # This tmp folder stands in for W:\LASER, so it has to be approved the same
    # way the real one is - an email may only pull from the shop's own shares.
    monkeypatch.setattr(job_intake_service, "APPROVED_SOURCE_ROOTS", (tmp_path / "W",))
    _dxf_with_text(shared / "Panel.dxf", "GEOMETRY")
    (shared / "Panel.pdf").write_bytes(b"%PDF-1.4\n")

    entry = job_intake_service.create_intake(
        "M50950", None, [], source="outlook", email_body=f"parts at {shared} thanks"
    )

    attachments = {a["filename"]: a for a in entry["attachments"]}
    assert attachments["Panel.dxf"]["in_place"] is True
    # The row points into the shared folder, not a copy under L:.
    assert Path(attachments["Panel.dxf"]["saved_path"]).parent == shared

    job_folder = Path(entry["job_folder"])
    assert not (job_folder / "Panel.dxf").exists(), "the DXF was copied after all"

    # The shared folder is left exactly as it was found.
    assert sorted(p.name for p in shared.iterdir()) == ["Panel.dxf", "Panel.pdf"]


def test_a_thickness_stated_in_the_email_is_used_when_nothing_else_says(
    tmp_path, monkeypatch, shop_csvs
) -> None:
    """It was extracted and then referenced nowhere.

    A source that is read but never consulted is worse than one that isn't read
    at all: the comment said it existed so it could dispute the others, so it
    looked handled.
    """
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    monkeypatch.setattr(
        job_intake_service, "MATERIAL_MEMORY_PATH", tmp_path / "fingerprints.json"
    )
    # The drawing names the material but no thickness - the common case.
    dxf = _dxf_with_text(tmp_path / "Bracket.dxf", "MATERIAL: ALUMINUM 5052")

    entry = job_intake_service.create_intake(
        "M50960", None, [dxf], email_body="please cut these in 1/4 THK aluminum 5052"
    )

    part = entry["material_qty"][0]
    assert part["material"] == "Aluminum 5052"
    assert part["thickness"] == 0.25, "the email's thickness was ignored"


def test_a_thickness_the_email_and_the_drawing_disagree_on_is_a_conflict(
    tmp_path, monkeypatch, shop_csvs
) -> None:
    """Every source that can choose a value must also be able to argue about
    it. Letting the email win silently is the same failure as letting any other
    source win silently - the metal gets cut either way."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    monkeypatch.setattr(
        job_intake_service, "MATERIAL_MEMORY_PATH", tmp_path / "fingerprints.json"
    )
    dxf = _dxf_with_text(
        tmp_path / "Bracket.dxf", "MATERIAL: ALUMINUM 5052", ".125 THK"
    )

    entry = job_intake_service.create_intake(
        "M50961", None, [dxf], email_body="cut these in 1/4 THK aluminum 5052"
    )

    part = entry["material_qty"][0]
    assert part["conflicts"].get("thickness"), "the disagreement went unnoticed"
    assert "the email" in part["conflicts"]["thickness"]

    # And it stops the import until someone settles it.
    part["material_confirmed"] = True
    part["qty_unknown"] = False
    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    assert "thickness" in str(excinfo.value)


def test_an_email_may_only_pull_from_the_shops_own_shares(
    tmp_path, monkeypatch, shop_csvs
) -> None:
    """An email body is untrusted text, and a path found in one gets listed and
    read. Without a boundary, any folder a message named - including one in a
    forwarded reply chain - could be pulled onto a job."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    monkeypatch.setattr(job_intake_service, "APPROVED_SOURCE_ROOTS", (tmp_path / "W",))

    elsewhere = tmp_path / "somebody-elses-folder"
    elsewhere.mkdir()
    _dxf_with_text(elsewhere / "NotOurs.dxf", "GEOMETRY")

    entry = job_intake_service.create_intake(
        "M50951", None, [], source="outlook", email_body=f"parts at {elsewhere} thanks"
    )

    assert entry["attachments"] == [], "read a folder it was not allowed to read"
    assert entry["ingested_from"] == []
    # Said so, rather than looking like the folder was empty.
    assert any("outside the folders" in note for note in entry["po_unmatched"])


@pytest.mark.parametrize(
    "label",
    [
        r"..\..\..\Windows\Temp",
        "../escape",
        r"sub\folder",
        "C:Rush",
        "..",
        "Rush?Plates",
        "CON",
        "nul.txt",
        "Rush Plates.",
        "Rush\tPlates",
        "x" * 65,
    ],
    ids=[
        "windows-traversal", "posix-traversal", "separator", "drive-relative",
        "dotdot", "invalid-char", "device-name", "device-name-with-extension",
        "trailing-dot", "control-character", "too-long",
    ],
)
def test_a_label_must_be_one_safe_folder_name(tmp_path, monkeypatch, label) -> None:
    """A label comes from an email subject or a typed field and is joined
    straight onto the job folder. Validated centrally, before any path is
    resolved from it - checking containment afterwards would catch the escape
    only after untrusted text had already decided the folder shape."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    with pytest.raises(JobIntakeError):
        job_intake_service.resolve_job_paths("M59919", label)


def test_an_ordinary_label_still_works(tmp_path, monkeypatch) -> None:
    """The rules must not reject the labels people actually use."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    for label in ("Rush Plates", "PFF PO-8527-001", "Repair Brackets (2nd)", "Console 1.5"):
        paths = job_intake_service.resolve_job_paths("M59919", label)
        assert paths.intake_dir.name == label

    # Surrounding whitespace is trimmed rather than refused - it is a typo, not
    # an attempt at anything, and Windows would have dropped it silently.
    assert job_intake_service.validate_label("  Rush Plates  ") == "Rush Plates"
    assert job_intake_service.validate_label("   ") is None
    assert job_intake_service.validate_label(None) is None


def test_creating_folders_refuses_a_target_outside_the_shop_root(
    tmp_path, monkeypatch
) -> None:
    """The containment check guarded delete and rename but not creation, which
    is the one that runs on every intake."""
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    paths = job_intake_service.resolve_job_paths("M59919", "Rush Plates")
    escaped = dataclasses.replace(paths, intake_dir=tmp_path / "elsewhere")

    with pytest.raises(JobIntakeError) as excinfo:
        job_intake_service.create_job_folders(escaped)
    assert "outside" in str(excinfo.value)
    assert not (tmp_path / "elsewhere").exists()


def test_a_placeholder_number_needs_a_label_and_parks_each_job_separately(
    tmp_path, monkeypatch
) -> None:
    """Work often arrives before its number does. A placeholder lets it be
    filed now, but several unrelated jobs must not pile into one folder while
    they wait, so each parks under its own label - typically the customer's PO
    number - and Rename Job moves it once the real number exists.
    """
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(
        job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", tmp_path / "registry.json"
    )
    dxf_a = _dxf_with_text(tmp_path / "A.dxf", "GEOMETRY")
    dxf_b = _dxf_with_text(tmp_path / "B.dxf", "GEOMETRY")

    assert job_intake_service.is_placeholder_job_number("M12345")
    assert job_intake_service.label_required_for("M12345")
    # A real number is only constrained once its folder exists.
    assert not job_intake_service.is_placeholder_job_number("M59919")

    with pytest.raises(JobIntakeError) as excinfo:
        job_intake_service.create_intake("M12345", None, [dxf_a])
    assert "placeholder" in str(excinfo.value)

    first = job_intake_service.create_intake("M12345", "PFF PO-8527-001", [dxf_a])
    second = job_intake_service.create_intake("M12345", "PFF PO-8600-002", [dxf_b])

    assert first["provisional"] is True
    assert Path(first["job_folder"]).name == "PFF PO-8527-001"
    assert Path(second["job_folder"]).name == "PFF PO-8600-002"
    # Same placeholder, different folders - nothing collided.
    assert first["job_folder"] != second["job_folder"]
    assert Path(first["job_folder"]).parent == Path(second["job_folder"]).parent

    # The prefix still picks the root, so an M placeholder lands under
    # M-FABRICATION like any other M job.
    assert "M-FABRICATION" in first["job_folder"]


def test_radan_import_refuses_while_radan_is_open_or_already_importing(tmp_path, monkeypatch) -> None:
    """Driving RADAN over COM while someone has it open can corrupt the project
    they are working in, and two imports writing one RPD is the same problem
    twice. Both checks are truck_nest_explorer's own, re-exported by its
    services module - not reimplemented here."""
    from types import SimpleNamespace

    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    paths = job_intake_service.resolve_job_paths("M59919", None)

    radan_open = SimpleNamespace(
        visible_radan_sessions=lambda: ((1234, "RADAN - somebody's job.rpd"),),
        radan_csv_import_lock_status=lambda _p: (False, "lock", None),
    )
    with pytest.raises(JobIntakeError) as excinfo:
        job_intake_service.assert_radan_is_safe_to_drive(radan_open, paths)
    assert "RADAN is open" in str(excinfo.value)
    assert "1234" in str(excinfo.value)

    already_importing = SimpleNamespace(
        visible_radan_sessions=lambda: (),
        radan_csv_import_lock_status=lambda _p: (True, tmp_path / "x.lock", 999),
    )
    with pytest.raises(JobIntakeError) as excinfo:
        job_intake_service.assert_radan_is_safe_to_drive(already_importing, paths)
    assert "already running" in str(excinfo.value)

    clear = SimpleNamespace(
        visible_radan_sessions=lambda: (),
        radan_csv_import_lock_status=lambda _p: (False, "lock", None),
    )
    job_intake_service.assert_radan_is_safe_to_drive(clear, paths)

    # A sibling app too old to expose the checks must not block the import.
    job_intake_service.assert_radan_is_safe_to_drive(SimpleNamespace(), paths)


# --- BOM conversion outcomes -------------------------------------------------


def _bom(tmp_path: Path, name: str = "BOM.csv") -> Path:
    """A file that _looks_like_bom_spreadsheet accepts."""
    path = tmp_path / name
    path.write_text("Part Number,Description,QTY\nA-1,PLATE 11GA,2\n", encoding="utf-8")
    return path


class InventorToRadanNeedsUi(RuntimeError):
    """Stands in for inventor_to_radan's own class of the same name.

    Deliberately defined here rather than imported: inline_runner execs that
    module under a private name and drops it again, so the class the converter
    actually raises is rebuilt on every call and is never the one this process
    would import. A same-named local class is therefore the faithful stand-in -
    and it is what catches a classifier written with isinstance, which would
    silently never match in production.
    """

    def __init__(self, missing_rules=(), missing_dxf_items=()):
        self.missing_rules = list(missing_rules)
        self.missing_dxf_items = list(missing_dxf_items)
        super().__init__("needs input")


def _runner_raising(exc: Exception):
    from types import SimpleNamespace

    def run_inline(*args, **kwargs):
        raise exc

    return SimpleNamespace(run_inline=run_inline)


def test_a_bom_needing_a_rule_is_named_and_blocks(tmp_path: Path, monkeypatch) -> None:
    """A missing CAM rule is recoverable in one specific way, so say which way.

    The old blanket handler returned an empty result, which was indistinguishable
    from "the BOM had no rows" - so the job carried on against the scraped
    prints, quietly demoting its own source of truth to a guess.
    """
    monkeypatch.setattr(
        job_intake_service,
        "_load_inventor_to_radan",
        lambda: _runner_raising(InventorToRadanNeedsUi(missing_rules=["PLATE 11GA", "TUBE 2X2"])),
    )
    conversion = job_intake_service.convert_bom_spreadsheet(_bom(tmp_path))

    assert conversion.outcome == job_intake_service.BOM_NEEDS_RULES
    assert conversion.blocked
    assert conversion.rows == ()
    # Names the descriptions and the way the shop adds a rule.
    assert "PLATE 11GA" in conversion.reason
    assert "inventor_to_radan" in conversion.reason


def test_a_bom_needing_a_dxf_classified_is_distinguished_from_one_needing_a_rule(
    tmp_path: Path, monkeypatch
) -> None:
    monkeypatch.setattr(
        job_intake_service,
        "_load_inventor_to_radan",
        lambda: _runner_raising(InventorToRadanNeedsUi(missing_dxf_items=[{"part": "A-1"}])),
    )
    conversion = job_intake_service.convert_bom_spreadsheet(_bom(tmp_path))

    assert conversion.outcome == job_intake_service.BOM_NEEDS_DXF_CLASSIFICATION
    assert conversion.blocked


@pytest.mark.parametrize(
    "exc, expected",
    [
        (ImportError("pandas is not installed"), "converter_unavailable"),
        (ValueError("no header row"), "unexpected_failure"),
    ],
    ids=["missing-dependency", "anything-else"],
)
def test_other_converter_failures_are_typed_rather_than_swallowed(
    tmp_path: Path, monkeypatch, exc, expected
) -> None:
    monkeypatch.setattr(
        job_intake_service, "_load_inventor_to_radan", lambda: _runner_raising(exc)
    )
    conversion = job_intake_service.convert_bom_spreadsheet(_bom(tmp_path))
    assert conversion.outcome == expected
    assert conversion.blocked
    assert str(exc) in conversion.reason


def test_a_missing_bom_is_not_a_failure(tmp_path: Path) -> None:
    """Nothing to convert is the ordinary case for a job filed from DXFs alone,
    and must not be confused with a BOM that would not convert."""
    conversion = job_intake_service.convert_bom_spreadsheet(tmp_path / "nope.csv")
    assert conversion.outcome == job_intake_service.BOM_NOT_ATTEMPTED
    assert not conversion.blocked


def test_a_bom_that_would_not_convert_blocks_the_import(tmp_path: Path) -> None:
    """The whole point of typing the outcome: the gate refuses, and says why."""
    entry = _entry_with_parts(tmp_path)
    entry["bom_blockers"] = [
        "BOM.csv uses 2 description(s) with no RADAN rule yet: PLATE 11GA, TUBE 2X2."
    ]

    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    message = str(excinfo.value)
    assert "STOP" in message
    assert "PLATE 11GA" in message
    assert "Nothing has been imported" in message

    # Cleared once the BOM converts, which is what re-applying it does.
    entry["bom_blockers"] = []
    assert len(build_import_csv_rows(entry)) == 1


def test_conflicting_sources_are_a_hard_stop(tmp_path: Path) -> None:
    """Two sources giving different answers isn't a note to read past. Whoever
    asked for the job has to say which is right - ranking one source over the
    other would just be picking a winner silently, and the metal gets cut
    either way."""
    entry = _entry_with_parts(tmp_path)
    part = entry["material_qty"][0]
    part["conflicts"] = {
        "material": "the CAM BOM says Mild Steel-A36, the print says Aluminum 5052"
    }
    part["resolved"] = {}

    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    message = str(excinfo.value)
    assert "STOP" in message
    # Both readings are named, so the user knows what to go and ask about.
    assert "Mild Steel-A36" in message and "Aluminum 5052" in message
    assert "whoever requested the job" in message

    # A human deciding that field is the only thing that clears it.
    part["resolved"] = {"material": True}
    assert len(build_import_csv_rows(entry)) == 1


def test_settling_the_material_does_not_release_a_disputed_quantity(tmp_path: Path) -> None:
    """The gate is per field. Ticking Verified on the material used to clear
    every conflict on the row, including a quantity two sources disagreed
    about - which is the one that decides how much metal gets cut."""
    entry = _entry_with_parts(tmp_path)
    part = entry["material_qty"][0]
    part["conflicts"] = {"quantity": "the CAM BOM says 12, the print says 1"}
    part["resolved"] = {"material": True, "material_confirmed": True}
    part["material_confirmed"] = True

    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    assert "quantity" in str(excinfo.value)

    part["resolved"]["quantity"] = True
    assert len(build_import_csv_rows(entry)) == 1


def test_build_import_csv_rows_refuses_an_unverified_material(tmp_path: Path) -> None:
    """A predicted material is a guess until someone checks it, and a wrong
    material is expensive - so the import is blocked rather than trusting it."""
    entry = _entry_with_parts(tmp_path)
    entry["material_qty"][0]["material_confirmed"] = False
    entry["material_qty"][0]["material_source_text"] = "MATL: CRS"

    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    message = str(excinfo.value)
    assert "confirm the material" in message
    # The customer's own wording is quoted so the user knows what to check.
    assert "MATL: CRS" in message

    entry["material_qty"][0]["material_confirmed"] = True
    assert len(build_import_csv_rows(entry)) == 1


def test_build_import_csv_rows_happy_path_and_write(tmp_path: Path) -> None:
    entry = _entry_with_parts(tmp_path)
    rows = build_import_csv_rows(entry)
    assert rows == [[str(tmp_path / "Clip-End.DXF"), "10", "Mild Steel-A36", "0.25", "in", "O2"]]

    csv_path = write_import_csv(rows, tmp_path / "out" / "import.csv")
    with csv_path.open(newline="", encoding="utf-8") as handle:
        assert list(csv.reader(handle)) == rows


def test_build_import_csv_rows_reports_all_problems(tmp_path: Path) -> None:
    entry = _entry_with_parts(tmp_path)
    entry["material_qty"][0]["material"] = ""
    entry["material_qty"].append(
        {"filename": "Missing.DXF", "material": "Aluminum 5052", "thickness": 0.12, "unit": "in", "qty": 2, "strategy": "", "material_confirmed": True}
    )
    entry["material_qty"].append(
        {"filename": "Clip-End.DXF", "material": "Aluminum 5052", "thickness": 0, "unit": "in", "qty": 2, "strategy": "", "material_confirmed": True}
    )
    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    message = str(excinfo.value)
    assert "pick a material" in message
    assert "Missing.DXF" in message
    assert "thickness" in message


# --- PO extraction -----------------------------------------------------------
# Synthetic PDFs are built with the same one-cell-per-line layout PyMuPDF
# produces for the real PFF PO template (verified against 5 real POs on L:).


def _write_po_pdf(tmp_path: Path, lines: list[str]) -> Path:
    fitz = pytest.importorskip("fitz")
    pdf_path = tmp_path / "po.pdf"
    doc = fitz.open()
    page = doc.new_page()
    y = 40.0
    for line in lines:
        page.insert_text((40, y), line)
        y += 14.0
    doc.save(str(pdf_path))
    doc.close()
    return pdf_path


PO_BODY = [
    "Date:",
    "PO Number:",
    "Date Required:",
    "Line",
    "Qty",
    "DESCRIPTION",
    "PRICING",
    "Subtotal",
    "1",
    "10",
    'Clip-End - 1/4" Mild Steel',
    "2",
    "36",
    'Clip-Mid - 1/4" Mild Steel',
    "3",
    "2",
    "End Cap D_2",
    "4",
    'ALL MATERIAL 1/8" MILD STEEL',
    "5",
    "6",
    "Sub-total",
    "0",
    "LASER ORDER",
    "July 21, 2026",
    "8665-001",
    "July 28, 2026",
]


def test_extract_po_hints_matches_lines_and_reports_unmatched(tmp_path: Path) -> None:
    pdf_path = _write_po_pdf(tmp_path, PO_BODY)
    hints = extract_po_hints(pdf_path, ["Clip-End", "Clip-Mid", "End Cap D", "End Cap D_2"])

    assert hints.po_number == "8665-001"
    assert hints.due_date == date(2026, 7, 28)
    assert hints.line_items["Clip-End"] == {"qty": 10, "raw_description": 'Clip-End - 1/4" Mild Steel'}
    assert hints.line_items["Clip-Mid"]["qty"] == 36
    # Longest-stem-first: D_2's row must not be claimed by "End Cap D".
    assert hints.line_items["End Cap D_2"]["qty"] == 2
    assert "End Cap D" not in hints.line_items
    # The order-wide material note surfaces as unmatched; footer labels don't.
    assert 'ALL MATERIAL 1/8" MILD STEEL' in hints.unmatched_lines
    assert not any("Sub-total" in line for line in hints.unmatched_lines)


def test_extract_po_hints_flags_po_lines_with_no_dxf(tmp_path: Path) -> None:
    pdf_path = _write_po_pdf(tmp_path, PO_BODY)
    hints = extract_po_hints(pdf_path, ["Clip-End"])
    assert 'Clip-Mid - 1/4" Mild Steel' in hints.unmatched_lines


def test_extract_po_hints_single_date_is_not_a_due_date(tmp_path: Path) -> None:
    # A lone date is the order date (Date Required said RUSH/ASAP or was
    # blank) - never claim it as the due date, but surface the urgency note.
    body = [line if line != "July 28, 2026" else "RUSH" for line in PO_BODY]
    hints = extract_po_hints(_write_po_pdf(tmp_path, body), ["Clip-End"])
    assert hints.due_date is None
    assert hints.due_note == "RUSH"
    assert hints.po_number == "8665-001"

    blank_body = [line for line in PO_BODY if line != "July 28, 2026"]
    blank_hints = extract_po_hints(_write_po_pdf(tmp_path, blank_body), ["Clip-End"])
    assert blank_hints.due_date is None
    assert blank_hints.due_note is None


def test_extract_po_hints_survives_non_pdf_garbage(tmp_path: Path) -> None:
    garbage = tmp_path / "not_really.pdf"
    garbage.write_bytes(b"this is not a pdf")
    hints = extract_po_hints(garbage, ["Clip-End"])
    assert hints.po_number is None
    assert hints.due_date is None
    assert hints.line_items == {}
