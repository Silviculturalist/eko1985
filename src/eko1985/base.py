"""Base classes for the Eko 1985 species parts."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, TYPE_CHECKING

from .enums import Trädslag
from .utils import qmd_cm

if TYPE_CHECKING:
    from .stand import EkoStand



class EvenAgedStand:
    """Minimal helpers shared by all stand components."""

    @staticmethod
    def getQMD(BA: float, stems: float) -> float:
        return qmd_cm(BA, stems)

    @staticmethod
    def getMAI(volume: float, total_age: float) -> float:
        return volume / total_age


@dataclass
class EkoStandPart(EvenAgedStand):
    """Common state for every species-specific cohort."""

    BA: float
    stems: float
    age: float
    trädslag: Trädslag
    stand: Optional["EkoStand"] = None
    QMDOtherSpecies: float = 0.0
    BAOtherSpecies: float = 0.0
    ba_quotient_acute_mortality: float = 0.0
    QMD: float = 0.0
    HK: float = 0.0

    def __post_init__(self) -> None:
        self.BA = float(self.BA)
        self.stems = float(self.stems)
        self.age = float(self.age)
        self.QMD = self.getQMD(self.BA, self.stems)

    def register_stand(self, stand: "EkoStand") -> None:
        self.stand = stand

    # Alias used in tests’ _snapshot()
    def volume_m3sk_ha(self, BA: float, QMD: float, age: float, stems: float, HK: float) -> float:
        return self.getVolume(BA=BA, QMD=QMD, age=age, stems=stems, HK=HK)


__all__ = ["EvenAgedStand", "EkoStandPart"]