from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Dict, List
from xml.etree import ElementTree as ET

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = REPO_ROOT / "Output2.xlsx"
EXPECTED_PATH = Path(__file__).resolve().parent / "data" / "output2_expected.json"

SHEET_REL_PATHS = {
    "General": "xl/worksheets/sheet1.xml",
    "Oversight": "xl/worksheets/sheet2.xml",
    "Site Variables": "xl/worksheets/sheet3.xml",
    "Biomass": "xl/worksheets/sheet4.xml",
}

NAMESPACE = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@pytest.fixture(scope="session")
def expected_output() -> Dict[str, List[Dict[str, str | None]]]:
    with EXPECTED_PATH.open(encoding="utf-8") as fh:
        return json.load(fh)


def _load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        xml = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(xml)
    out: List[str] = []
    for si in root:
        text = "".join(
            node.text or ""
            for node in si.iter()
            if node.tag.endswith("}t")
        )
        out.append(text)
    return out


def _cell_value(cell: ET.Element, shared_strings: List[str]) -> str | None:
    value_node = cell.find("main:v", NAMESPACE)
    if value_node is None:
        return None

    if cell.attrib.get("t") == "s":
        return shared_strings[int(value_node.text)]

    return value_node.text


def _load_sheet_rows(sheet_rel_path: str) -> List[Dict[str, str | None]]:
    with zipfile.ZipFile(OUTPUT_PATH) as zf:
        shared_strings = _load_shared_strings(zf)
        xml = zf.read(sheet_rel_path)

    root = ET.fromstring(xml)
    rows: List[Dict[str, str | None]] = []

    for row in root.find("main:sheetData", NAMESPACE):
        row_data: Dict[str, str | None] = {}
        for cell in row:
            coord = cell.attrib["r"]
            column = "".join(filter(str.isalpha, coord))
            row_data[column] = _cell_value(cell, shared_strings)
        rows.append(row_data)

    return rows


@pytest.mark.parametrize("sheet", sorted(SHEET_REL_PATHS))
def test_output2_sheet_matches_expected(sheet: str, expected_output: Dict[str, List[Dict[str, str | None]]]) -> None:
    actual_rows = _load_sheet_rows(SHEET_REL_PATHS[sheet])
    assert sheet in expected_output, f"Missing expected data for sheet '{sheet}'"
    assert actual_rows == expected_output[sheet]


def _general_table(rows: List[Dict[str, str | None]]) -> Dict[tuple[str, str, str], Dict[str, str | None]]:
    table: Dict[tuple[str, str, str], Dict[str, str | None]] = {}
    period = ""
    phase = ""

    for row in rows:
        if "A" in row and row["A"] is not None:
            period = row["A"]
        if "B" in row and row["B"] is not None:
            phase = row["B"]
        species = row.get("C")
        if not species:
            continue

        key = (period, phase, species)
        table[key] = {
            "stems": row.get("G"),
            "basal_area": row.get("H"),
            "qmd": row.get("I"),
            "volume": row.get("J"),
            "stems_removed": row.get("L"),
            "basal_area_removed": row.get("M"),
            "qmd_removed": row.get("N"),
            "volume_removed": row.get("O"),
            "net_increment": row.get("Q"),
            "mean_increment": row.get("R"),
        }

    return table


def test_general_sheet_key_values(expected_output: Dict[str, List[Dict[str, str | None]]]) -> None:
    rows = expected_output["General"]
    table = _general_table(rows)

    assert table[("3", "Gallring", "Tall")]["stems_removed"] == "94"
    assert table[("3", "Gallring", "Tall")]["volume_removed"] == "28"
    assert table[("9", "Tillväxt", "Gran")]["volume"] == "81"
    assert table[("9", "Tillväxt", "Öv.löv")]["stems"] == "145"
    assert table[("0", "Start", "Summa")]["volume"] == "365"