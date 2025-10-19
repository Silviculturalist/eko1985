"""Stand container and compatibility helpers."""
from __future__ import annotations

from math import pi
import warnings

from .base import EkoStandPart, EvenAgedStand
from .site import EkoStandSite
from .species import EkoBeech, EkoOak
from .utils import safe_sum


class EkoStand(EvenAgedStand):
    """Aggregate stand container combining cohorts with their site."""
    def __init__(self, parts, site: EkoStandSite):
        if not all(isinstance(p, EkoStandPart) for p in parts):
            raise ValueError("All StandParts must be EkoStandPart objects!")
        self.parts = parts           # Swedish API used in tests
        self.Parts = self.parts      # Back-compat alias
        self.Site = site

        # species-region sanity (as in your original warnings)
        if any(isinstance(p, EkoBeech) for p in self.parts) and self.Site.region not in ("Central", "South"):
            warnings.warn("Setting Beech stand outside of southern Sweden!")
        if any(isinstance(p, EkoOak) for p in self.parts) and self.Site.region not in ("Central", "South"):
            warnings.warn("Setting Oak stand outside of southern Sweden!")

        for p in self.parts:
            p.register_stand(self)

        # initialize all metrics once
        self._assign_current_state_metrics()

    @staticmethod
    def _volume_for(part: EkoStandPart, BA: float, QMD: float, age: float, stems: float, HK: float) -> float:
        if BA is None or BA <= 0 or stems is None or stems <= 0:
            return 0.0
        try:
            return part.getVolume(BA=BA, QMD=QMD, age=age, stems=stems, HK=HK)
        except ValueError:
            return 0.0

    # ------------------------------------------------------------------
    # NEW helper: competition & interaction (HK) on the *current* state
    # ------------------------------------------------------------------
    def _competition_metrics(self) -> None:
        """Compute per-part competition metrics from the current net state."""
        # ensure own QMD is current
        for p in self.parts:
            p.QMD = self.getQMD(p.BA, p.stems)
        # competition from others
        for p in self.parts:
            BA_other = safe_sum(q.BA for q in self.parts if q is not p)
            N_other = safe_sum(q.stems for q in self.parts if q is not p)
            p.BAOtherSpecies = BA_other
            p.QMDOtherSpecies = self.getQMD(BA_other, N_other)
            denom = p.QMD if p.QMD > 0 else 1e-9
            p.HK = (p.QMDOtherSpecies / denom) * BA_other

    # ------------------------------------------------------------------
    # NEW helper: fully assign all CURRENT state metrics
    # ------------------------------------------------------------------
    def _assign_current_state_metrics(self) -> None:
        """
        Recompute everything derived from the current net state:
          - per-part: QMD, competition metrics, volume (p.VOL)
          - stand totals: StandBA, StandStems, StandVOL
        """
        self._competition_metrics()
        for p in self.parts:
            p.VOL = self._volume_for(p, p.BA, p.QMD, p.age, p.stems, p.HK)

        self.StandBA = safe_sum(p.BA for p in self.parts)
        self.StandStems = safe_sum(p.stems for p in self.parts)
        self.StandVOL = safe_sum(p.VOL for p in self.parts)

    # Back-compat alias used by tests’ _snapshot()
    def _refresh_competition_vars(self) -> None:
        self._assign_current_state_metrics()

    # Thinning: constant-QMD removal; refresh metrics afterward
    def thin(self, removals: dict):
        """
        removals: dict keyed by part.trädslag (enum) or species string; value = BA (m²/ha) to remove.
        Stems removed computed at current QMD (constant‑QMD removal).
        """
        for p in self.parts:
            key_variants = (p.trädslag, getattr(p.trädslag, "value", None), str(getattr(p.trädslag, "value", "")))
            BA_out = 0.0
            for k in key_variants:
                if k in removals:
                    BA_out = float(removals[k])
                    break
            BA_out = max(0.0, min(BA_out, p.BA))
            if BA_out > 0.0 and p.QMD > 0:
                stems_out = BA_out / (pi * (p.QMD / 200.0) ** 2)
                p.BA -= BA_out
                p.stems = max(0.0, p.stems - stems_out)
        self._assign_current_state_metrics()

    # Five-year growth step (thesis order-of-ops). Returns dict with N1/BA1/QMD1/VOL1.
    def grow(self, years: int = 5, apply_mortality: bool = True):
        # 0) Start-of-period: initialize competition/interaction and totals
        self._assign_current_state_metrics()

        # Start-of-period snapshot volumes for bookkeeping
        for p in self.parts:
            p.VOL0 = self._volume_for(p, p.BA, p.QMD, p.age, p.stems, p.HK)

        # 1) gross BAI + 2) natural mortality
        for p in self.parts:
            if apply_mortality:
                BAQ_crowd, QMD_dead_crowd, BAQ_other, QMD_dead_other = p.getMortality(increment=years)
            else:
                BAQ_crowd = BAQ_other = 0.0
                QMD_dead_crowd = 0.9 * (p.QMD or 0.0)
                QMD_dead_other = (p.QMD or 0.0)

            # gross basal area increment using correct mortality quotients as inputs
            p.getBAI5(
                ba_quotient_chronic_mortality=BAQ_crowd,
                ba_quotient_acute_mortality=BAQ_other,
            )

            # End-of-period before removals
            p.BA1 = p.BA + p.BAI5
            p.stems1 = p.stems  # no ingrowth modeled here
            p.age2 = p.age + years
            p.QMD1 = self.getQMD(p.BA1, p.stems1)

            # Natural mortality based on start-of-period BA
            p.BA_Mortality = (BAQ_crowd + BAQ_other) * p.BA
            if p.QMD > 0:
                stems_other = p.BA * BAQ_other / (pi * ((QMD_dead_other / 200.0) ** 2))
                stems_crowd = p.BA * BAQ_crowd / (pi * ((QMD_dead_crowd / 200.0) ** 2))
            else:
                stems_other = stems_crowd = 0.0
            p.stems_Mortality = max(0.0, stems_other + stems_crowd)

        # 5a) competition + volume at end pre‑removal
        for p in self.parts:
            BA_other = safe_sum(q.BA1 for q in self.parts if q is not p)
            N_other = safe_sum(q.stems1 for q in self.parts if q is not p)
            QMD_other = self.getQMD(BA_other, N_other)
            p.HK1 = (QMD_other / (p.QMD1 if p.QMD1 > 0 else 1e-9)) * BA_other

        for p in self.parts:
            p.VOL1 = self._volume_for(p, p.BA1, p.QMD1, p.age2, p.stems1, p.HK1)

        # 5b) after natural mortality (no thinning inside grow(); do thinning between periods)
        for p in self.parts:
            p.BA2 = max(0.0, p.BA1 - p.BA_Mortality)
            p.stems2 = max(0.0, p.stems1 - p.stems_Mortality)
            p.QMD2 = self.getQMD(p.BA2, p.stems2)

        for p in self.parts:
            BA_other = safe_sum(q.BA2 for q in self.parts if q is not p)
            N_other = safe_sum(q.stems2 for q in self.parts if q is not p)
            QMD_other = self.getQMD(BA_other, N_other)
            p.HK2 = (QMD_other / (p.QMD2 if p.QMD2 > 0 else 1e-9)) * BA_other

        for p in self.parts:
            p.VOL2 = self._volume_for(p, p.BA2, p.QMD2, p.age2, p.stems2, p.HK2)

        # 6) Commit net end-of-period state; return shape expected by tests
        period = {}
        for p in self.parts:
            p.gross_volume_increment = p.VOL1 - p.VOL0
            p.BA, p.stems, p.QMD, p.age, p.VOL = p.BA2, p.stems2, p.QMD2, p.age2, p.VOL2

            key = p.trädslag.value  # "Tall", "Gran", "Björk", ...
            period.setdefault(key, []).append(
                {
                    "N1": p.stems,
                    "BA1": p.BA,
                    "QMD1": p.QMD,
                    "VOL1": p.VOL,
                    # optional provenance (ignored by your asserts)
                    "N0": p.stems1,
                    "BA0": p.BA1,
                    "QMD0": p.QMD1,
                    "VOL0": p.VOL1,
                }
            )

        # final housekeeping for the new state
        self._assign_current_state_metrics()
        return period

    # legacy wrapper used by some old calls
    def grow5(self, mortality=True):
        return self.grow(years=5, apply_mortality=mortality)

def install_eko_compat():
    """
    Make total-stand metrics always available and safe:
      • EkoStand.StandBA and EkoStand.StandStems become lazy properties with setters.
      • QMDs are precomputed (if needed) before the first volume assignment pass.
    Call this once AFTER your Eko* classes are defined and BEFORE constructing any EkoStand.
    """
    g = globals()
    if 'EkoStand' not in g:
        raise RuntimeError("Define EkoStand/EkoStandSite and species classes first.")

    EkoStand = g['EkoStand']

    if getattr(EkoStand, "_eko_forward_safe", False):
        return  # already installed

    # ---- Lazy totals: work even before constructor sets them explicitly
    def _parts(self):
        return getattr(self, 'Parts', getattr(self, 'parts', [])) or []

    def _get_StandBA(self):
        # Use cached value if later set by the model; otherwise compute on the fly
        if hasattr(self, "_StandBA"):
            return self._StandBA
        return float(sum((getattr(p, "BA", 0.0) or 0.0) for p in _parts(self)))

    def _set_StandBA(self, v):
        self._StandBA = float(v)

    def _get_StandStems(self):
        if hasattr(self, "_StandStems"):
            return self._StandStems
        return float(sum((getattr(p, "stems", 0.0) or 0.0) for p in _parts(self)))

    def _set_StandStems(self, v):
        self._StandStems = float(v)

    # Attach as data-descriptor properties (safe to assign later)
    EkoStand.StandBA = property(_get_StandBA, _set_StandBA, doc="Total basal area (m²/ha), lazy-safe.")
    EkoStand.StandStems = property(_get_StandStems, _set_StandStems, doc="Total stems (/ha), lazy-safe.")

    # ---- Guard the first metrics pass: ensure QMD exists before any volume call
    # If your class already does this, the wrapper is a no-op overhead.
    orig = getattr(EkoStand, "_assign_current_state_metrics", None)
    if callable(orig):
        def _assign_current_state_metrics_wrapped(self, *a, **kw):
            parts = _parts(self)
            # Ensure QMD is available before any formula may try to use it
            for p in parts:
                if (getattr(p, "QMD", None) in (None, 0)) and (getattr(p, "BA", 0) > 0) and (getattr(p, "stems", 0) > 0):
                    try:
                        p.QMD = self.getQMD(p.BA, p.stems)
                    except Exception:
                        pass
            return orig(self, *a, **kw)
        EkoStand._assign_current_state_metrics = _assign_current_state_metrics_wrapped

    EkoStand._eko_forward_safe = True


__all__ = ["EkoStand", "install_eko_compat"]