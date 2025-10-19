"""Excel parsing helpers used by the Eko 1985 tooling."""
from __future__ import annotations

from pathlib import Path
import os

import pandas as pd


def _to_num(x):
    try:
        if pd.isna(x): 
            return None
    except Exception:
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
    return None if s == 'nan' else s

# ----------------------------- #
# Sheet parsers
# ----------------------------- #

def _parse_site_variables(xlsx_path: str) -> dict:
    """
    Reads the 'Site Variables' sheet and extracts:
      - latitude, altitude, region, soil_moisture_code (1=dry, 3=mesic, 5=wet),
        vegetation_code (1 herbs/grass, 13 bilberry, 14 cowberry ... minimal map here),
      - H100 per species from 'Ståndortsindex, dm' (converted to meters: dm/10).
    """
    df = pd.read_excel(Path(xlsx_path), sheet_name="Site Variables", header=None)

    def find(label):
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                if str(df.iat[i, j]).strip().lower() == label.lower():
                    return (i, j)
        return None

    # Geographic
    lat_pos = find('Latitud')
    alt_pos = find('Altitud')
    omr_pos = find('Område')
    lat = _to_num(df.iat[lat_pos[0]+1, lat_pos[1]]) if lat_pos else None
    alt = _to_num(df.iat[alt_pos[0]+1, alt_pos[1]]) if alt_pos else None
    region_raw = _to_str(df.iat[omr_pos[0]+1, omr_pos[1]]) if omr_pos else None
    region_map = {
        'Syd': 'South', 'Södra': 'South', 'SÖDRA': 'South', 'Södra Sverige': 'South',
        'Mellan': 'Central', 'Centrala': 'Central', 'Central': 'Central',
        'Nord': 'North', 'Norra': 'North'
    }
    # region_raw may be None; use an empty-string fallback so the key is always a str for type checkers
    region = region_map.get(region_raw or "", None)

    # Soil moisture (very simple mapping)
    torr_pos = find('Torr')
    vat_pos = find('Våt')
    soil_code = 3
    if torr_pos and _to_str(df.iat[torr_pos[0]+1, torr_pos[1]]) in ('Ja', 'True', '1'):
        soil_code = 1
    if vat_pos and _to_str(df.iat[vat_pos[0]+1, vat_pos[1]]) in ('Ja', 'True', '1'):
        soil_code = 5

    # Vegetation (very simple mapping)
    ort_pos = find('Ört/gräs')
    blabar_pos = find('Blåbär/lingon')
    veg_code = None
    if ort_pos and _to_str(df.iat[ort_pos[0]+1, ort_pos[1]]) in ('Ja', 'True', '1'):
        veg_code = 1
    if blabar_pos and _to_str(df.iat[blabar_pos[0]+1, blabar_pos[1]]) in ('Ja', 'True', '1'):
        veg_code = 13

    # H100 site indices (dm → m)
    si_label = find('Ståndortsindex, dm')
    H100 = {}
    if si_label:
        species_row = df.iloc[si_label[0]+1].tolist()
        values_row = df.iloc[si_label[0]+2].tolist()
        for name, val in zip(species_row, values_row):
            if pd.isna(name) or pd.isna(val):
                continue
            H100[str(name)] = _to_num(val) / 10.0 if _to_num(val) is not None else None

    return {
        'latitude': lat,
        'altitude_m': alt,
        'region': region,
        'soil_moisture_code': soil_code,
        'vegetation_code': veg_code,
        'H100': H100
    }

def _parse_general_sheet(xlsx_path: str) -> list[dict]:
    """
    Reads 'General' and returns a list of events of the form:
      {'period': int, 'type': 'Start'|'Tillväxt'|'Gallring',
       'species': { 'Tall': {...}, 'Gran': {...}, ... }}

    Under each species we include:
      - ages: total_age, bh_age, h_top_m
      - 'after' state (Stamantal st/ha, Grundyta m2/ha, Dg cm, Volym m3sk/ha)
      - extraction (Uttag/självgallring), growth (Årlig tillväxt), and mortality flags
    """
    df = pd.read_excel(Path(xlsx_path), sheet_name="General", header=None)
    events_idx = [i for i in range(len(df)) if df.iloc[i, 1] in ('Start', 'Tillväxt', 'Gallring')]
    species_order = ['Tall', 'Gran', 'Björk', 'Bok', 'Ek', 'Öv.löv']

    events = []
    for ei, i in enumerate(events_idx):
        typ = df.iloc[i, 1]
        period = df.iloc[i, 0]
        try:
            period = int(period)
        except Exception:
            period = ei if typ == 'Start' else (events[-1]['period'] + 1 if events else 0)

        # species rows for a block can start 1 row above the event label in some exports (e.g., 'Tall' above 'Start')
        next_boundary = next((idx for idx in events_idx if idx > i), len(df))
        species_block = {}
        k = max(0, i - 1)
        while k < next_boundary:
            sp = df.iloc[k, 2]
            if sp in species_order:
                species_block[sp] = {
                    'total_age': _to_num(df.iloc[k, 3]),
                    'bh_age': _to_num(df.iloc[k, 4]),
                    'h_top_m': _to_num(df.iloc[k, 5]),
                    'after': {
                        'N_stems_ha': _to_num(df.iloc[k, 6]),
                        'BA_m2_ha': _to_num(df.iloc[k, 7]),
                        'QMD_cm': _to_num(df.iloc[k, 8]),
                        'VOL_m3sk_ha': _to_num(df.iloc[k, 9]),
                    },
                    'extraction': {
                        'N_stems_ha': _to_num(df.iloc[k, 11]),
                        'BA_m2_ha': _to_num(df.iloc[k, 12]),
                        'QMD_cm': _to_num(df.iloc[k, 13]),
                        'VOL_m3sk_ha': _to_num(df.iloc[k, 14]),
                    },
                    'growth': {
                        'lopande_m3sk_ha': _to_num(df.iloc[k, 16]),
                        'medel_m3sk_ha': _to_num(df.iloc[k, 17]),
                    },
                    'mortality': {
                        'slow_BA_frac': _to_num(df.iloc[k, 18]),
                        'fast_BA_frac': _to_num(df.iloc[k, 19]),
                    },
                    'flags': {
                        'nygallara': _to_str(df.iloc[k, 20]),
                        'gallrad_nagongang': _to_str(df.iloc[k, 21]),
                        'gallringshistorik': _to_str(df.iloc[k, 22]),
                    },
                }
            k += 1

        events.append({'period': period, 'type': typ, 'species': species_block})

    return events

def excel_to_json(xlsx_path: str) -> dict:
    """
    High-level: parse the Excel and return a single, tidy structure.
    """
    return {
        'source_file': os.path.basename(xlsx_path),
        'site': _parse_site_variables(xlsx_path),
        'events': _parse_general_sheet(xlsx_path),
    }



__all__ = ["excel_to_json"]