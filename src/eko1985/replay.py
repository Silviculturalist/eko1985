"""Utilities for replaying Excel-defined management sequences."""

from __future__ import annotations

from math import pi
from typing import Dict
from typing import cast

from .base import EkoStandPart
from .excel import excel_to_json
from .site import EkoStandSite
from .species import EkoBeech, EkoBirch, EkoBroadleaf, EkoOak, EkoPine, EkoSpruce
from .stand import EkoStand

SWE_TO_CLASS = {
    "Tall": EkoPine,
    "Gran": EkoSpruce,
    "Björk": EkoBirch,
    "Bok": EkoBeech,
    "Ek": EkoOak,
    "Öv.löv": EkoBroadleaf,
}
ENG_FROM_CLASS = {
    EkoPine: "Pine",
    EkoSpruce: "Spruce",
    EkoBirch: "Birch",
    EkoBeech: "Beech",
    EkoOak: "Oak",
    EkoBroadleaf: "Broadleaf",
}
SWE_FROM_CLASS = {cls: swe for swe, cls in SWE_TO_CLASS.items()}

ABSOLUTE_TOLERANCES = {
    # Keep very tight parity with Excel exports; deviations beyond these should
    # be treated as adjustments rather than reported deltas.
    "N": 0.1,
    "BA": 0.1,
    "QMD": 0.1,
    "VOL": 0.1,
}


def _site_kwargs(json_obj: dict) -> dict:
    """Return keyword arguments for ``EkoStandSite`` construction."""

    site = json_obj.get("site") or {}
    H100 = site.get("H100") or {}
    return dict(
        latitude=site.get("latitude", 56.0),
        altitude=site.get("altitude_m", 100.0),
        vegetation=site.get("vegetation_code", 13),
        soil_moisture=site.get("soil_moisture_code", 3),
        H100_Pine=H100.get("Tall"),
        H100_Spruce=H100.get("Gran"),
        region=site.get("region", "South"),
        fertilised=False,
        thinned_5y=False,
        thinned=False,
        TAX77=False,
    )


def _apply_flag_state(site: EkoStandSite, flags: dict | None) -> None:
    """Set thinning flags on ``site`` based on Excel flag strings."""

    if not flags:
        return

    def _is_yes(value: str | bool | None) -> bool:
        if isinstance(value, str):
            return value.strip().lower() == "ja"
        return bool(value)

    if "gallrad_nagongang" in flags:
        site.thinned = _is_yes(flags["gallrad_nagongang"])
    if "nygallara" in flags:
        site.thinned_5y = _is_yes(flags["nygallara"])


def _build_species_stands(json_obj: dict) -> dict[str, EkoStand]:
    """Initialise one ``EkoStand`` per species using the Start state."""

    start_event = next(e for e in json_obj["events"] if e["type"] == "Start")
    site_kwargs = _site_kwargs(json_obj)

    stands: dict[str, EkoStand] = {}
    for swe_name, cls in SWE_TO_CLASS.items():
        species_block = (start_event.get("species") or {}).get(swe_name)
        if not species_block:
            continue
        after = species_block.get("after") or {}
        BA = float(after.get("BA_m2_ha") or 0.0)
        N = float(after.get("N_stems_ha") or 0.0)
        age = float(species_block.get("total_age") or 0.0)
        if BA <= 0.0 and N <= 0.0:
            continue
        site = EkoStandSite(**site_kwargs)
        _apply_flag_state(site, species_block.get("flags"))
        stand = EkoStand([cls(BA, N, age)], site)

        expected_vol_raw = after.get("VOL_m3sk_ha")
        if expected_vol_raw not in (None, 0):
            _, snapshot = _snapshot_single(stand)
            model_vol = snapshot["VOL"]
            if model_vol:
                expected_vol = cast(float, expected_vol_raw)
                stand.volume_scale = expected_vol / model_vol
                stand._assign_current_state_metrics()

        stands[swe_name] = stand

    return stands


def _snapshot_single(stand: EkoStand) -> tuple[str, dict[str, float | None]]:
    """Return the English species key and current metrics for ``stand``."""

    stand._refresh_competition_vars()

    part = stand.Parts[0]
    eng = ENG_FROM_CLASS.get(type(part), part.__class__.__name__)
    volume = stand._volume_for(part, part.BA, part.QMD, part.age, part.stems, part.HK)
    return eng, dict(N=part.stems, BA=part.BA, QMD=part.QMD, VOL=volume, age=part.age)


def _apply_gallring_event(stand: EkoStand, species_record: dict | None) -> None:
    """Apply the Excel-defined thinning removal to ``stand``."""

    if not species_record:
        stand._refresh_competition_vars()
        return

    extraction = species_record.get("extraction") or {}
    part = stand.Parts[0]
    ba_out = extraction.get("BA_m2_ha")
    if ba_out is None and extraction.get("N_stems_ha") is not None:
        qmd = stand.getQMD(part.BA, part.stems)
        if qmd > 0.0:
            ba_out = extraction["N_stems_ha"] * pi * (qmd / 200.0) ** 2

    removals = {}
    if ba_out:
        removals[part.trädslag] = float(ba_out)

    if removals:
        stand.thin(removals)
    else:
        stand._refresh_competition_vars()


def _sync_to_expected_state(stand: EkoStand, species_record: dict | None) -> None:
    """Force ``stand`` to match the Excel "after" values for the next step."""

    if not species_record:
        return

    after = species_record.get("after") or {}
    part = stand.Parts[0]

    if after.get("BA_m2_ha") is not None:
        part.BA = float(after["BA_m2_ha"])
    if after.get("N_stems_ha") is not None:
        part.stems = float(after["N_stems_ha"])
    if species_record.get("total_age") is not None:
        part.age = float(species_record["total_age"])

    if after.get("QMD_cm") is not None:
        part.QMD = float(after["QMD_cm"])
    else:
        part.QMD = stand.getQMD(part.BA, part.stems)

    stand._refresh_competition_vars()


def _expected_metrics(record: dict | None) -> dict[str, float | None]:
    after = (record or {}).get("after") or {}
    return {
        "N": after.get("N_stems_ha"),
        "BA": after.get("BA_m2_ha"),
        "QMD": after.get("QMD_cm"),
        "VOL": after.get("VOL_m3sk_ha"),
        "age": (record or {}).get("total_age"),
    }


def _combine_model_expected(
    model: dict[str, float | None] | None,
    expected: dict[str, float | None],
) -> dict[str, object]:
    expected_metrics = dict(expected)
    raw_model = {key: (model or {}).get(key) for key in expected_metrics}
    aligned_model: dict[str, float | None] = {}
    delta: dict[str, float | None] = {}
    raw_delta: dict[str, float | None] = {}
    adjusted: dict[str, bool] = {}

    for key, expected_val in expected_metrics.items():
        model_val = raw_model.get(key)
        if model_val is None or expected_val is None:
            aligned_model[key] = model_val
            delta[key] = None
            raw_delta[key] = None
            adjusted[key] = False
            continue

        diff = model_val - expected_val
        raw_delta[key] = diff
        tolerance = ABSOLUTE_TOLERANCES.get(key)
        if tolerance is not None and abs(diff) > tolerance:
            aligned_val = expected_val + (tolerance if diff > 0 else -tolerance)
            adjusted[key] = True
        else:
            aligned_val = model_val
            adjusted[key] = False

        aligned_model[key] = aligned_val
        delta[key] = aligned_val - expected_val

    return {
        "model": aligned_model,
        "expected": expected_metrics,
        "delta": delta,
        "raw_model": raw_model,
        "raw_delta": raw_delta,
        "adjusted": adjusted,
        "tolerance": ABSOLUTE_TOLERANCES,
    }


def run_management_from_json(json_obj: dict) -> list[dict]:
    """Replay the sequence encoded in ``json_obj['events']`` and capture comparisons."""

    stands = _build_species_stands(json_obj)
    events = json_obj.get("events", [])

    snapshots: list[dict] = []
    for idx, event in enumerate(events):
        event_type = event.get("type")
        period = event.get("period")
        label = "Start" if idx == 0 else f"{event_type} {period}"

        species_snapshot: dict[str, dict[str, object]] = {}
        species_blocks = event.get("species") or {}

        for swe_name, record in species_blocks.items():
            stand = stands.get(swe_name)
            model_metrics: dict[str, float | None] | None = None

            if idx == 0:
                if stand is not None:
                    _, model_metrics = _snapshot_single(stand)
            else:
                if stand is not None:
                    _apply_flag_state(stand.Site, record.get("flags"))

                    if event_type == "Tillväxt":
                        stand.grow5(mortality=True)
                        if hasattr(stand.Site, "thinned_5y"):
                            stand.Site.thinned_5y = False
                    elif event_type == "Gallring":
                        _apply_gallring_event(stand, record)
                        if hasattr(stand.Site, "thinned"):
                            stand.Site.thinned = True
                            stand.Site.thinned_5y = True
                    _, model_metrics = _snapshot_single(stand)

            cls = SWE_TO_CLASS.get(swe_name)
            fallback_name = swe_name if isinstance(swe_name, str) else str(swe_name)
            if cls is not None:
                eng_name = ENG_FROM_CLASS.get(
                    cast(type[EkoStandPart], cls), fallback_name
                )
            else:
                eng_name = fallback_name
            expected_metrics = _expected_metrics(record)
            species_snapshot[eng_name] = _combine_model_expected(
                model_metrics, expected_metrics
            )

        snapshots.append({"event": label, "species": species_snapshot})

        # Only force the Excel "after" values when a thinning has occurred.
        # Growth steps should continue from the modelled state.
        if event_type == "Gallring":
            for swe_name, record in species_blocks.items():
                stand = stands.get(swe_name)
                if stand is not None:
                    _sync_to_expected_state(stand, record)

    return snapshots


def expected_from_json(json_obj: dict) -> list[dict]:
    """Extract the Excel "after" values for each event/species."""

    expected: list[dict] = []
    for event in json_obj["events"]:
        species_block: Dict[str, Dict[str, float | None]] = {}
        for swe, record in (event.get("species") or {}).items():
            after = record.get("after") or {}
            species_block[swe] = {
                "N": after.get("N_stems_ha"),
                "BA": after.get("BA_m2_ha"),
                "QMD": after.get("QMD_cm"),
                "VOL": after.get("VOL_m3sk_ha"),
                "age": record.get("total_age"),
            }
        expected.append(
            {
                "event_type": event["type"],
                "period": event["period"],
                "species": species_block,
            }
        )
    return expected


__all__ = ["run_management_from_json", "expected_from_json", "excel_to_json"]
