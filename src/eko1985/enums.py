"""Enumerations used across the Eko 1985 model."""
from __future__ import annotations

from enum import Enum


class RegionSE(Enum):
    """Supported Swedish regions."""

    NORRA = "North"
    MELLERSTA = "Central"
    SÖDRA = "South"


class VegetationsKod(Enum):
    """Subset of vegetation codes used in the tests."""

    ÖRTER_1 = 1
    BLÅBÄR = 13


class MarkfuktighetKod(Enum):
    """Subset of soil moisture codes used in the tests."""

    TORR_ELLER_MYCKET_TORR = 1
    FRISK = 3
    VÅT_ELLER_MYCKET_VÅT = 5


class Trädslag(Enum):
    """Species handled by the model."""

    TALL = "Tall"
    GRAN = "Gran"
    BJÖRK = "Björk"
    BOK = "Bok"
    EK = "Ek"
    ÖV_LÖV = "Öv.löv"


__all__ = [
    "RegionSE",
    "VegetationsKod",
    "MarkfuktighetKod",
    "Trädslag",
]