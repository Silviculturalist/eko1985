
"""Utility helpers shared across the model."""
from __future__ import annotations

from math import pi, sqrt
from typing import Iterable


def qmd_cm(BA_m2_per_ha: float, stems_per_ha: float) -> float:
    """Quadratic mean diameter (cm) given basal area and stem count."""

    if BA_m2_per_ha <= 0.0 or stems_per_ha <= 0.0:
        return 0.0
    return sqrt(BA_m2_per_ha * 40000.0 / (pi * stems_per_ha))


def safe_sum(iterable: Iterable[float]) -> float:
    """Sum ``iterable`` converting members to floats."""

    total = 0.0
    for value in iterable:
        total += float(value)
    return total


__all__ = ["qmd_cm", "safe_sum"]