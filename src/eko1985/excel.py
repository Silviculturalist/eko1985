"""Excel parsing helpers used by the Eko 1985 tooling."""

from __future__ import annotations

from pathlib import Path
import math
import os
import posixpath
from typing import Iterable, List, Sequence, overload
import xml.etree.ElementTree as ET
from zipfile import ZipFile


class _SheetRow:
    """Lightweight row proxy exposing a tiny pandas-compatible surface."""

    def __init__(self, values: Sequence[object], width: int) -> None:
        self._values = list(values)
        self._width = width

    def __getitem__(self, index: int) -> object:
        if index < 0 or index >= self._width:
            raise IndexError(index)
        if index < len(self._values):
            return self._values[index]
        return None

    def tolist(self) -> list[object]:
        return [self[i] for i in range(self._width)]


class _ILocAccessor:
    def __init__(self, sheet: "_Sheet") -> None:
        self._sheet = sheet

    @overload
    def __getitem__(self, key: tuple[int, int]) -> object:
        ...

    @overload
    def __getitem__(self, key: int) -> _SheetRow:
        ...

    def __getitem__(self, key: int | tuple[int, int]) -> object | _SheetRow:  # type: ignore[override]
        if isinstance(key, tuple):
            row, col = key
            if isinstance(row, slice) or isinstance(
                col, slice
            ):  # pragma: no cover - sanity
                raise TypeError("slice access is not supported")
            return self._sheet._get_value(int(row), int(col))
        return self._sheet._get_row(int(key))


class _Sheet:
    """Very small subset of the :class:`pandas.DataFrame` API we rely on."""

    def __init__(self, rows: Iterable[Sequence[object]]) -> None:
        materialized_rows: List[List[object]] = [list(row) for row in rows]
        self._rows = materialized_rows
        self._width = max((len(row) for row in materialized_rows), default=0)
        self.iloc = _ILocAccessor(self)

    @property
    def shape(self) -> tuple[int, int]:
        return (len(self._rows), self._width)

    def __len__(self) -> int:
        return len(self._rows)

    def _get_row(self, index: int) -> _SheetRow:
        if index < 0 or index >= len(self._rows):
            raise IndexError(index)
        return _SheetRow(self._rows[index], self._width)

    def _get_value(self, row: int, col: int) -> object:
        if row < 0 or row >= len(self._rows):
            raise IndexError(row)
        if col < 0:
            raise IndexError(col)
        if col < len(self._rows[row]):
            return self._rows[row][col]
        return None

    def iat(self, row: int, col: int) -> object:
        return self._get_value(row, col)


def _is_na(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    return False


def _to_num(x):
    try:
        if _is_na(x):
            return None
    except Exception:  # pragma: no cover - mirrors original defensive behaviour
        pass
    try:
        return float(x)
    except Exception:
        try:
            # Some exports use comma decimal separator
            s = str(x).replace(",", ".")
            return float(s)
        except Exception:
            return None


def _to_str(x):
    s = str(x)
    return None if s == "nan" else s


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _column_index_from_ref(ref: str) -> int:
    col = 0
    for ch in ref:
        if ch.isalpha():
            col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
        else:
            break
    return max(col - 1, 0)


def _read_shared_strings(zf: ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    tree = ET.fromstring(data)
    ns = {"main": _MAIN_NS}
    strings: list[str] = []
    for si in tree.findall("main:si", ns):
        parts = [t.text or "" for t in si.findall(".//main:t", ns)]
        strings.append("".join(parts))
    return strings


def _read_sheet_targets(zf: ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_map = {}
    for rel in relationships.findall(f"{{{_PKG_REL_NS}}}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            if target.startswith("/"):
                rel_map[rel_id] = target.lstrip("/")
            else:
                rel_map[rel_id] = posixpath.normpath(posixpath.join("xl", target))

    sheets = {}
    ns = {"main": _MAIN_NS, "rel": _REL_NS}
    for sheet in workbook.findall("main:sheets/main:sheet", ns):
        name = sheet.get("name")
        rel_id = sheet.get(f"{{{_REL_NS}}}id")
        if name and rel_id and rel_id in rel_map:
            sheets[name] = rel_map[rel_id]
    return sheets


def _read_cell_value(
    cell: ET.Element, shared_strings: list[str], ns: dict[str, str]
) -> object:
    cell_type = cell.get("t")
    if cell_type == "s":
        value = cell.find("main:v", ns)
        if value is None or value.text is None:
            return None
        idx = int(value.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else None
    if cell_type == "b":
        value = cell.find("main:v", ns)
        return (
            value.text == "1" if value is not None and value.text is not None else None
        )
    if cell_type == "inlineStr":
        parts = [t.text or "" for t in cell.findall(".//main:t", ns)]
        return "".join(parts)
    value = cell.find("main:v", ns)
    if value is None or value.text is None:
        return None
    if cell_type == "str":
        return value.text
    try:
        return float(value.text)
    except (TypeError, ValueError):
        return value.text


def _extract_sheet_rows(data: bytes, shared_strings: list[str]) -> list[list[object]]:
    ns = {"main": _MAIN_NS}
    sheet = ET.fromstring(data)
    rows: list[list[object]] = []
    for row in sheet.findall("main:sheetData/main:row", ns):
        current: list[object] = []
        last_col = -1
        for cell in row.findall("main:c", ns):
            ref = cell.get("r")
            col_idx = _column_index_from_ref(ref or "") if ref else last_col + 1
            while len(current) <= col_idx:
                current.append(None)
            current[col_idx] = _read_cell_value(cell, shared_strings, ns)
            last_col = col_idx
        rows.append(current)
    return rows


def _load_sheet(xlsx_path: Path, sheet_name: str) -> _Sheet:
    with ZipFile(xlsx_path) as zf:
        shared_strings = _read_shared_strings(zf)
        targets = _read_sheet_targets(zf)
        try:
            sheet_path = targets[sheet_name]
        except KeyError as exc:  # pragma: no cover - invalid input guard
            raise KeyError(f"sheet '{sheet_name}' not found") from exc
        rows = _extract_sheet_rows(zf.read(sheet_path), shared_strings)
    return _Sheet(rows)


# ----------------------------- #
# Sheet parsers
# ----------------------------- #


def _parse_site_variables(sheet: _Sheet) -> dict:
    """
    Reads the 'Site Variables' sheet and extracts:
      - latitude, altitude, region, soil_moisture_code (1=dry, 3=mesic, 5=wet),
        vegetation_code (1 herbs/grass, 13 bilberry, 14 cowberry ... minimal map here),
      - H100 per species from 'Ståndortsindex, dm' (converted to meters: dm/10).
    """

    def find(label):
        for i in range(sheet.shape[0]):
            for j in range(sheet.shape[1]):
                if str(sheet.iat(i, j)).strip().lower() == label.lower():
                    return (i, j)
        return None

    # Geographic
    lat_pos = find("Latitud")
    alt_pos = find("Altitud")
    omr_pos = find("Område")
    lat = _to_num(sheet.iat(lat_pos[0] + 1, lat_pos[1])) if lat_pos else None
    alt = _to_num(sheet.iat(alt_pos[0] + 1, alt_pos[1])) if alt_pos else None
    region_raw = _to_str(sheet.iat(omr_pos[0] + 1, omr_pos[1])) if omr_pos else None
    region_map = {
        "Syd": "South",
        "Södra": "South",
        "SÖDRA": "South",
        "Södra Sverige": "South",
        "Mellan": "Central",
        "Centrala": "Central",
        "Central": "Central",
        "Nord": "North",
        "Norra": "North",
    }
    # region_raw may be None; use an empty-string fallback so the key is always a str for type checkers
    region = region_map.get(region_raw or "", None)

    # Soil moisture (very simple mapping)
    torr_pos = find("Torr")
    vat_pos = find("Våt")
    soil_code = 3
    if torr_pos and _to_str(sheet.iat(torr_pos[0] + 1, torr_pos[1])) in (
        "Ja",
        "True",
        "1",
    ):
        soil_code = 1
    if vat_pos and _to_str(sheet.iat(vat_pos[0] + 1, vat_pos[1])) in (
        "Ja",
        "True",
        "1",
    ):
        soil_code = 5

    # Vegetation (very simple mapping)
    ort_pos = find("Ört/gräs")
    blabar_pos = find("Blåbär/lingon")
    veg_code = None
    if ort_pos and _to_str(sheet.iat(ort_pos[0] + 1, ort_pos[1])) in (
        "Ja",
        "True",
        "1",
    ):
        veg_code = 1
    if blabar_pos and _to_str(sheet.iat(blabar_pos[0] + 1, blabar_pos[1])) in (
        "Ja",
        "True",
        "1",
    ):
        veg_code = 13

    # H100 site indices (dm → m)
    si_label = find("Ståndortsindex, dm")
    H100 = {}
    if si_label:
        species_row = sheet.iloc[si_label[0] + 1].tolist()
        values_row = sheet.iloc[si_label[0] + 2].tolist()
        for name, val in zip(species_row, values_row):
            if _is_na(name) or _is_na(val):
                continue
            parsed = _to_num(val)
            H100[str(name)] = parsed / 10.0 if parsed is not None else None

    return {
        "latitude": lat,
        "altitude_m": alt,
        "region": region,
        "soil_moisture_code": soil_code,
        "vegetation_code": veg_code,
        "H100": H100,
    }


def _parse_general_sheet(sheet: _Sheet) -> list[dict]:
    """
    Reads 'General' and returns a list of events of the form:
      {'period': int, 'type': 'Start'|'Tillväxt'|'Gallring',
       'species': { 'Tall': {...}, 'Gran': {...}, ... }}

    Under each species we include:
      - ages: total_age, bh_age, h_top_m
      - 'after' state (Stamantal st/ha, Grundyta m2/ha, Dg cm, Volym m3sk/ha)
      - extraction (Uttag/självgallring), growth (Årlig tillväxt), and mortality flags
    """
    events_idx = [
        i
        for i in range(len(sheet))
        if sheet.iloc[i, 1] in ("Start", "Tillväxt", "Gallring")
    ]
    species_order = ["Tall", "Gran", "Björk", "Bok", "Ek", "Öv.löv"]

    events: list[dict] = []
    for ei, i in enumerate(events_idx):
        typ = sheet.iloc[i, 1]
        period_raw = sheet.iloc[i, 0]
        if isinstance(period_raw, (int, float, str)):
            try:
                period = int(period_raw)
            except Exception:
                period = (
                    ei if typ == "Start" else (events[-1]["period"] + 1 if events else 0)
                )
        else:
            period = ei if typ == "Start" else (events[-1]["period"] + 1 if events else 0)

        # species rows for a block can start 1 row above the event label in some exports (e.g., 'Tall' above 'Start')
        next_boundary = next((idx for idx in events_idx if idx > i), len(sheet))
        species_block = {}
        k = max(0, i - 1)
        while k < next_boundary:
            sp = sheet.iloc[k, 2]
            if sp in species_order:
                species_block[sp] = {
                    "total_age": _to_num(sheet.iloc[k, 3]),
                    "bh_age": _to_num(sheet.iloc[k, 4]),
                    "h_top_m": _to_num(sheet.iloc[k, 5]),
                    "after": {
                        "N_stems_ha": _to_num(sheet.iloc[k, 6]),
                        "BA_m2_ha": _to_num(sheet.iloc[k, 7]),
                        "QMD_cm": _to_num(sheet.iloc[k, 8]),
                        "VOL_m3sk_ha": _to_num(sheet.iloc[k, 9]),
                    },
                    "extraction": {
                        "N_stems_ha": _to_num(sheet.iloc[k, 11]),
                        "BA_m2_ha": _to_num(sheet.iloc[k, 12]),
                        "QMD_cm": _to_num(sheet.iloc[k, 13]),
                        "VOL_m3sk_ha": _to_num(sheet.iloc[k, 14]),
                    },
                    "growth": {
                        "lopande_m3sk_ha": _to_num(sheet.iloc[k, 16]),
                        "medel_m3sk_ha": _to_num(sheet.iloc[k, 17]),
                    },
                    "mortality": {
                        "slow_BA_frac": _to_num(sheet.iloc[k, 18]),
                        "fast_BA_frac": _to_num(sheet.iloc[k, 19]),
                    },
                    "flags": {
                        "nygallara": _to_str(sheet.iloc[k, 20]),
                        "gallrad_nagongang": _to_str(sheet.iloc[k, 21]),
                        "gallringshistorik": _to_str(sheet.iloc[k, 22]),
                    },
                }
            k += 1

        events.append({"period": period, "type": typ, "species": species_block})

    return events


def _parse_oversight_extractions(
    sheet: _Sheet,
) -> list[dict[str, dict[str, float | None]]]:
    """
    Reads the 'Oversight' sheet and returns per-Gallring removal blocks.

    Structure observed in the supplied workbooks:
      - A block headed with ``Gallring N`` followed by one row per species with
        age, stems, basal area, and volume removed.
    """

    species_order = ["Tall", "Gran", "Björk", "Bok", "Ek", "Öv.löv"]
    gallrings: list[dict[str, dict[str, float | None]]] = []
    current: dict[str, dict[str, float | None]] | None = None

    for i in range(len(sheet)):
        label = sheet.iloc[i, 0]
        if isinstance(label, str) and label.startswith("Gallring"):
            current = {}
            gallrings.append(current)
            continue

        if current is None:
            continue

        if not isinstance(label, str):
            continue

        if label in species_order:
            # Columns (by observation): [species, age, N, BA, VOL, ...]
            age = _to_num(sheet.iloc[i, 1])
            stems = _to_num(sheet.iloc[i, 2])
            ba = _to_num(sheet.iloc[i, 3])
            vol = _to_num(sheet.iloc[i, 4])
            current[label] = {
                "total_age": age,
                "N_stems_ha": stems,
                "BA_m2_ha": ba,
                "VOL_m3sk_ha": vol,
            }

    return gallrings


def _resolve_workbook_path(path: Path) -> Path:
    if path.exists():
        return path
    repo_root = Path(__file__).resolve().parents[2]
    candidates = [
        repo_root / path.name,
        repo_root / "assets" / path.name,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(path)


def excel_to_json(xlsx_path: str) -> dict:
    """
    High-level: parse the Excel and return a single, tidy structure.
    """
    original_path = Path(xlsx_path)
    path = _resolve_workbook_path(original_path)
    site_sheet = _load_sheet(path, "Site Variables")
    general_sheet = _load_sheet(path, "General")
    # Prefer explicit thinning removals from 'Oversight' when present; fall
    # back to the 'General' extraction columns otherwise.
    try:
        oversight_sheet = _load_sheet(path, "Oversight")
        oversight_extractions = _parse_oversight_extractions(oversight_sheet)
    except Exception:
        oversight_extractions = []

    events = _parse_general_sheet(general_sheet)

    gallring_idx = 0
    for ev in events:
        if ev["type"] != "Gallring":
            continue
        if gallring_idx >= len(oversight_extractions):
            gallring_idx += 1
            continue
        block = oversight_extractions[gallring_idx]
        for swe_name, removal in block.items():
            species_rec = ev["species"].get(swe_name)
            if species_rec is None:
                continue
            # Replace extraction with Oversight's explicit removal values
            species_rec["extraction"] = {
                "N_stems_ha": removal.get("N_stems_ha"),
                "BA_m2_ha": removal.get("BA_m2_ha"),
                "QMD_cm": None,
                "VOL_m3sk_ha": removal.get("VOL_m3sk_ha"),
            }
            # Prefer Oversight age for the removal if provided
            if removal.get("total_age") is not None:
                species_rec["total_age"] = removal["total_age"]
        gallring_idx += 1

    return {
        "source_file": os.path.basename(xlsx_path),
        "site": _parse_site_variables(site_sheet),
        "events": events,
    }


__all__ = ["excel_to_json"]
