"""Utilities for replaying Excel-defined management sequences."""
from __future__ import annotations

from typing import Dict

from .excel import excel_to_json
from .site import EkoStandSite
from .species import EkoBeech, EkoBirch, EkoBroadleaf, EkoOak, EkoPine, EkoSpruce
from .stand import EkoStand

SWE_TO_CLASS = {
    'Tall': EkoPine,
    'Gran': EkoSpruce,
    'Björk': EkoBirch,
    'Bok': EkoBeech,
    'Ek': EkoOak,
    'Öv.löv': EkoBroadleaf,
}
ENG_FROM_CLASS = {
    EkoPine: 'Pine',
    EkoSpruce: 'Spruce',
    EkoBirch: 'Birch',
    EkoBeech: 'Beech',
    EkoOak: 'Oak',
    EkoBroadleaf: 'Broadleaf',
}
SWE_FROM_CLASS = {cls: swe for swe, cls in SWE_TO_CLASS.items()}


def _build_stand_from_json(json_obj: dict) -> tuple[EkoStandSite, EkoStand]:
    """Build ``EkoStandSite`` and ``EkoStand`` instances from parsed JSON."""

    site = json_obj['site'] or {}
    H100 = site.get('H100', {})

    site_kwargs = dict(
        latitude=site.get('latitude', 56.0),
        altitude=site.get('altitude_m', 100.0),
        vegetation=site.get('vegetation_code', 13),
        soil_moisture=site.get('soil_moisture_code', 3),
        H100_Pine=H100.get('Tall', None),
        H100_Spruce=H100.get('Gran', None),
        region=site.get('region', 'South'),
        fertilised=False,
        thinned_5y=False,
        thinned=False,
        TAX77=False,
    )

    stand_site = EkoStandSite(**site_kwargs)

    start_event = next(e for e in json_obj['events'] if e['type'] == 'Start')
    parts = []
    for swe_name, cls in SWE_TO_CLASS.items():
        if swe_name in start_event['species']:
            rec = start_event['species'][swe_name]
            after = rec.get('after') or {}
            BA = after.get('BA_m2_ha') or 0.0
            N = after.get('N_stems_ha') or 0.0
            age = rec.get('total_age') or 0.0
            if BA > 0 or N > 0:
                parts.append(cls(BA, N, age))

    stand = EkoStand(parts, stand_site)
    return stand_site, stand


def _snapshot(stand: EkoStand, label: str | None = None) -> dict:
    """Capture current per-species values using the model's own volume function."""

    for part in stand.Parts:
        part.QMD = stand.getQMD(part.BA, part.stems)

    stand._refresh_competition_vars()

    snap = {'event': label, 'species': {}}
    for part in stand.Parts:
        volume = stand._volume_for(part, part.BA, part.QMD, part.age, part.stems, part.HK)
        eng = ENG_FROM_CLASS.get(type(part), part.__class__.__name__)
        snap['species'][eng] = dict(N=part.stems, BA=part.BA, QMD=part.QMD, VOL=volume, age=part.age)
    return snap


def _thin_to_match_after_state(stand: EkoStand, event_species_block: dict) -> None:
    """Instantly set BA/N to the "after" values for a Gallring event and refresh metrics."""

    for part in stand.Parts:
        swe_key = SWE_FROM_CLASS.get(type(part))
        if swe_key and swe_key in event_species_block:
            after = (event_species_block[swe_key] or {}).get('after') or {}
            if after.get('BA_m2_ha') is not None:
                part.BA = float(after['BA_m2_ha'])
            if after.get('N_stems_ha') is not None:
                part.stems = float(after['N_stems_ha'])
            part.QMD = stand.getQMD(part.BA, part.stems)

    stand._refresh_competition_vars()
    for part in stand.Parts:
        part.QMD = stand.getQMD(part.BA, part.stems)
        _ = stand._volume_for(part, part.BA, part.QMD, part.age, part.stems, part.HK)

    if hasattr(stand, 'Site'):
        setattr(stand.Site, 'thinned', True)
        setattr(stand.Site, 'thinned_5y', True)


def run_management_from_json(json_obj: dict) -> list[dict]:
    """Replay the sequence encoded in ``json_obj['events']`` using the model."""

    _, stand = _build_stand_from_json(json_obj)

    snapshots = []
    snapshots.append(_snapshot(stand, label="Start"))

    for event in json_obj['events'][1:]:
        if event['type'] == 'Tillväxt':
            stand.grow5(mortality=True)
            if hasattr(stand.Site, 'thinned_5y'):
                stand.Site.thinned_5y = False
            snapshots.append(_snapshot(stand, label=f"Tillväxt {event['period']}"))
        elif event['type'] == 'Gallring':
            _thin_to_match_after_state(stand, event['species'])
            snapshots.append(_snapshot(stand, label=f"Gallring {event['period']}"))
        else:
            snapshots.append(_snapshot(stand, label=f"{event['type']} {event['period']}"))

    return snapshots


def expected_from_json(json_obj: dict) -> list[dict]:
    """Extract the Excel "after" values for each event/species."""

    expected: list[dict] = []
    for event in json_obj['events']:
        species_block: Dict[str, Dict[str, float | None]] = {}
        for swe, record in (event.get('species') or {}).items():
            after = record.get('after') or {}
            species_block[swe] = {
                'N': after.get('N_stems_ha'),
                'BA': after.get('BA_m2_ha'),
                'QMD': after.get('QMD_cm'),
                'VOL': after.get('VOL_m3sk_ha'),
                'age': record.get('total_age'),
            }
        expected.append({'event_type': event['type'], 'period': event['period'], 'species': species_block})
    return expected


__all__ = ["run_management_from_json", "expected_from_json", "excel_to_json"]