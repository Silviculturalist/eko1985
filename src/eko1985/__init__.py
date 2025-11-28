"""Python package wrapping the original 1985 EKO notebook."""

from .base import EkoStandPart, EvenAgedStand
from .enums import MarkfuktighetKod, RegionSE, Trädslag, VegetationsKod
from .site import EkoStandSite
from .species import EkoBeech, EkoBirch, EkoBroadleaf, EkoOak, EkoPine, EkoSpruce
from .stand import EkoStand, install_eko_compat

__all__ = [
    "EkoStandPart",
    "EvenAgedStand",
    "MarkfuktighetKod",
    "RegionSE",
    "Trädslag",
    "VegetationsKod",
    "EkoStandSite",
    "EkoBeech",
    "EkoBirch",
    "EkoBroadleaf",
    "EkoOak",
    "EkoPine",
    "EkoSpruce",
    "EkoStand",
    "install_eko_compat",
]

install_eko_compat()
