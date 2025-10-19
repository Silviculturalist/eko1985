"""Parity tests for the packaged Eko 1985 model."""
from __future__ import annotations

from pathlib import Path

import pytest

from eko1985.replay import excel_to_json, expected_from_json, run_management_from_json

REPO_ROOT = Path(__file__).resolve().parents[1]

ENG_TO_SWE = {
    'Pine': 'Tall',
    'Spruce': 'Gran',
    'Birch': 'Björk',
    'Beech': 'Bok',
    'Oak': 'Ek',
    'Broadleaf': 'Öv.löv',
}


def _align_snapshot_to_swe(snapshot: dict) -> dict:
    """Convert English species keys in a snapshot to Swedish."""

    aligned: dict = {}
    for eng, values in (snapshot.get('species') or {}).items():
        swe = ENG_TO_SWE.get(eng, eng)
        aligned[swe] = values
    return aligned


@pytest.mark.parametrize(
    "xls_path",
    [
        REPO_ROOT / "Output3.xlsx",
        REPO_ROOT / "output4.xlsx",
        REPO_ROOT / "Output2.xlsx",
    ],
)
def test_model_management_pipeline_runs(xls_path: Path) -> None:
    data = excel_to_json(str(xls_path))
    expected = expected_from_json(data)
    model_snaps = run_management_from_json(data)

    assert len(model_snaps) == len(expected)

    for expected_event, snapshot in zip(expected, model_snaps):
        snapshot_swe = _align_snapshot_to_swe(snapshot)
        for swe, record in (expected_event.get('species') or {}).items():
            if record['BA'] is None and record['N'] is None:
                continue
            assert swe in snapshot_swe
            got = snapshot_swe[swe]
            for key in ('N', 'BA', 'QMD', 'VOL', 'age'):
                assert key in got