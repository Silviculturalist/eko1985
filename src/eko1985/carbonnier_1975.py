from dataclasses import dataclass
from typing import Sequence
import bisect
import numpy as np


@dataclass
class CarbonnierHeightModel:
    """
    Carbonnier (1975) height development model for European beech.
    
    The model is defined by:
        h_j = a_j + τ * b_j
    
    where:
        h_j    = top height (m) at total age j (years)
        a_j    = age-specific base height
        b_j    = age-specific modifier
        τ      = site quality parameter derived from H100 (height at age 100)
    
    Parameters
    ----------
    ages : list of float
        Ages (years) at which a_j and b_j are tabulated (5, 10, ..., 130).
    a_vals : list of float
        Sequence of a_j coefficients.
    b_vals : list of float
        Sequence of b_j coefficients (corrected order).
    """

    ages: Sequence[float]
    a_vals: Sequence[float]
    b_vals: Sequence[float]

    # ------------------------------------------------------------------
    # τ (tau) calculation
    # ------------------------------------------------------------------
    def _tau_from_SI(self, site_index: float, si_age: float = 100.0) -> float:
        """
        Compute τ (site quality parameter) from top height at age si_age.

        τ = (SI - a_si_age) / b_si_age
        """
        try:
            idx = self.ages.index(si_age)
        except ValueError:
            raise ValueError(
                f"si_age={si_age} not found in ages grid {self.ages}"
            )

        a_si = self.a_vals[idx]
        b_si = self.b_vals[idx]

        if b_si == 0:
            raise ZeroDivisionError("b coefficient at si_age is zero; cannot compute tau.")

        return (site_index - a_si) / b_si

    # ------------------------------------------------------------------
    # Height for a given age and τ
    # ------------------------------------------------------------------
    def _height_at_age(self, age: float, tau: float, interpolate: bool = True) -> float:
        """
        Compute height at arbitrary age for a given τ.

        If age is tabulated exactly, the formula h = a + τ b is used.
        Otherwise, linear interpolation is applied between the nearest
        age classes.
        """
        # Exact match
        if age in self.ages:
            i = self.ages.index(age)
            return self.a_vals[i] + tau * self.b_vals[i]

        if not interpolate:
            raise ValueError("Age not tabulated and interpolate=False.")

        # Locate surrounding ages
        i = bisect.bisect_left(self.ages, age)

        if i == 0 or i == len(self.ages):
            raise ValueError(
                f"Age {age} outside coefficient range [{self.ages[0]}, {self.ages[-1]}]."
            )

        # Neighbors
        age_lo, age_hi = self.ages[i - 1], self.ages[i]
        h_lo = self.a_vals[i - 1] + tau * self.b_vals[i - 1]
        h_hi = self.a_vals[i] + tau * self.b_vals[i]

        # Linear interpolation
        w = (age - age_lo) / (age_hi - age_lo)
        return h_lo + w * (h_hi - h_lo)

    # ------------------------------------------------------------------
    # Height curve given site index (H100)
    # ------------------------------------------------------------------
    def height_from_SI(self, age: float, site_index: float, si_age: float = 100.0) -> float:
        """
        Compute height at age for a stand with given site index (H100).
        """
        tau = self._tau_from_SI(site_index, si_age)
        return self._height_at_age(age, tau, interpolate=True)
