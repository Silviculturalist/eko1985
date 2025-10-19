"""Site description for the Eko 1985 stand model."""
from __future__ import annotations

from math import exp, log
from typing import Any
import warnings

from .enums import MarkfuktighetKod, RegionSE, VegetationsKod


class EkoStandSite:
    """Represents the site variables required by the model."""

    def __init__(
        self,
        # English
        latitude: float | None = None,
        altitude: float | None = None,
        vegetation: Any | None = None,
        soil_moisture: Any | None = None,
        H100_Spruce: float | None = None,
        region: str | RegionSE | None = None,
        H100_Pine: float | None = None,
        fertilised: bool = False,
        thinned_5y: bool = False,
        thinned: bool = False,
        TAX77: bool = False,
        # Swedish (aliases used in the notebook)
        latitud: float | None = None,
        höjd_möh: float | None = None,
        vegetation_: VegetationsKod | None = None,
        markfuktighet: MarkfuktighetKod | None = None,
        h100_gran: float | None = None,
        h100_tall: float | None = None,
        gödslad: bool | None = None,
        gallrad_senaste_5år: bool | None = None,
        gallrad: bool | None = None,
        tax77: bool | None = None,
        klimat_zon: str | None = None,
    ) -> None:
        # prefer Swedish names when provided; coerce numeric site vars to floats with safe defaults
        if latitud is not None:
            self.latitude = float(latitud)
        elif latitude is not None:
            self.latitude = float(latitude)
        else:
            self.latitude = 0.0

        if höjd_möh is not None:
            self.altitude = float(höjd_möh)
        elif altitude is not None:
            self.altitude = float(altitude)
        else:
            self.altitude = 0.0

        self.fertilised = bool(gödslad) if gödslad is not None else bool(fertilised)
        self.thinned_5y = bool(gallrad_senaste_5år) if gallrad_senaste_5år is not None else bool(thinned_5y)
        self.thinned = bool(gallrad) if gallrad is not None else bool(thinned)
        self.TAX77 = bool(tax77) if tax77 is not None else bool(TAX77)
        self.klimat_zon = klimat_zon

        # region normalization
        if isinstance(region, RegionSE):
            self.region = region.value
        elif isinstance(region, str) and region in ("North", "Central", "South"):
            self.region = region
        elif isinstance(region, str) and region in ("NORRA", "MELLERSTA", "SÖDRA"):
            self.region = {"NORRA": "North", "MELLERSTA": "Central", "SÖDRA": "South"}[region]
        else:
            self.region = RegionSE.SÖDRA.value if region is None else str(region)

        # site index (H100) handling; tests may pass one or both in Swedish
        if h100_gran is not None:
            H100_Spruce = h100_gran
        if h100_tall is not None:
            H100_Pine = h100_tall

        if H100_Spruce is None and H100_Pine is None:
            raise ValueError("At least one of h100_gran or h100_tall must be provided.")

        if H100_Spruce is None:
            self.H100_Pine = H100_Pine
            self.H100_Spruce = self.__Leijon_Pine_to_Spruce(H100_Pine)
        elif H100_Pine is None:
            self.H100_Spruce = H100_Spruce
            self.H100_Pine = self.__Leijon_Spruce_to_Pine(H100_Spruce)
        else:
            self.H100_Spruce = H100_Spruce
            self.H100_Pine = H100_Pine

        # vegetation & field layer flags + vegcode
        veg_in = vegetation
        if isinstance(vegetation, VegetationsKod):
            veg_in = vegetation.value
        self._set_fieldlayer_and_vegcode(veg_in, self.latitude)

        # soil moisture flags
        sm = soil_moisture
        if isinstance(markfuktighet, MarkfuktighetKod):
            sm = markfuktighet.value
        self.DrySoil = sm == 1
        self.WetSoil = sm == 5

    def _set_fieldlayer_and_vegcode(self, vegetation_code: int | None, latitude: float | None) -> None:
        # Set FieldLayer flags
        if vegetation_code in (13, 14):  # bilberry/cowberry
            self.Bilberry_or_Cowberry = True
            self.HerbsGrassesNoFieldLayer = False
        elif vegetation_code in (1, 2, 3, 4, 5, 6, 8, 9) or (vegetation_code == 7 and (latitude or 0) < 60):
            self.Bilberry_or_Cowberry = False
            self.HerbsGrassesNoFieldLayer = True
        else:
            self.Bilberry_or_Cowberry = False
            self.HerbsGrassesNoFieldLayer = False

        # vegcode mapping (as in the original switch)
        mapping = {
            1: 4,
            2: 2.5,
            3: 2,
            4: 3,
            5: 2.5,
            6: 2,
            7: 3,
            8: 2.5,
            9: 1.5,
            10: -3,
            11: -3,
            12: 1,
            13: 0,
            14: -0.5,
            15: -3,
            16: -5,
            17: -0.5,
            18: -1,
        }
        if isinstance(vegetation_code, int):
            self.vegcode = mapping.get(vegetation_code, 0)
        else:
            self.vegcode = 0

    # --- Leijon conversions (unchanged) ---
    @staticmethod
    def __Leijon_Pine_to_Spruce(H100_Pine: float | None) -> float:
        if H100_Pine is None:
            return 0.0
        if H100_Pine < 8 or H100_Pine > 30:
            warnings.warn("SI Pine may be outside underlying material")
        return exp(-0.9596 * log(H100_Pine * 10) + 0.01171 * (H100_Pine * 10) + 7.9209) / 10

    @staticmethod
    def __Leijon_Spruce_to_Pine(H100_Spruce: float | None) -> float:
        if H100_Spruce is None:
            return 0.0
        if H100_Spruce < 8 or H100_Spruce > 33:
            warnings.warn("SI Spruce may be outside underlying material.")
        return exp(1.6967 * log(H100_Spruce * 10) - 0.005179 * (H100_Spruce * 10) - 2.5397) / 10


__all__ = ["EkoStandSite"]