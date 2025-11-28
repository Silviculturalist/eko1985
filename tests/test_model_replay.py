"""Parity tests for the packaged Eko 1985 model."""
from __future__ import annotations

import json
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
SWE_TO_ENG = {swe: eng for eng, swe in ENG_TO_SWE.items()}

ABS_TOLERANCES = {
    'N': 0.1,
    'BA': 0.1,
    'QMD': 0.1,
    'VOL': 0.1,
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
def test_model_management_pipeline_runs(xls_path: Path, artifacts_dir: Path) -> None:
    data = excel_to_json(str(xls_path))
    expected = expected_from_json(data)
    model_snaps = run_management_from_json(data)

    workbook_stem = xls_path.stem
    expected_path = artifacts_dir / f"{workbook_stem}_expected.json"
    model_path = artifacts_dir / f"{workbook_stem}_model.json"

    expected_path.write_text(
        json.dumps(expected, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    model_path.write_text(
        json.dumps(model_snaps, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )

    assert len(model_snaps) == len(expected)

    for expected_event, snapshot in zip(expected, model_snaps):
        for swe, record in (expected_event.get('species') or {}).items():
            eng = SWE_TO_ENG.get(swe, swe)
            assert eng in snapshot['species']
            comparison = snapshot['species'][eng]

            model_values = comparison['model']
            expected_values = comparison['expected']
            deltas = comparison['delta']

            for key in ('N', 'BA', 'QMD', 'VOL', 'age'):
                assert expected_values[key] == record[key]

                model_value = model_values.get(key)
                expected_value = expected_values.get(key)
                delta_value = deltas.get(key)

                if model_value is None or expected_value is None:
                    assert delta_value is None
                else:
                    # Delta should match the reported model - expected
                    assert delta_value == pytest.approx(model_value - expected_value, abs=1e-9)
                    # And respect the per-field absolute tolerance
                    tol = ABS_TOLERANCES.get(key)
                    if tol is not None:
                        assert abs(delta_value) <= tol + 1e-9
