"""Stand container and compatibility helpers."""
from __future__ import annotations

from math import pi
import warnings

from .base import EkoStandPart, EvenAgedStand
from .site import EkoStandSite
from .species import EkoBeech, EkoOak

from math import pi

def _safe_sum(iterable):
    total = 0.0
    for v in iterable:
        total += float(v)
    return total


class EkoStand(EvenAgedStand):
    """
    parts: list[EkoStandPart]
    site:  EkoStandSite
    """
    def __init__(self, parts, site: EkoStandSite):
        if not all(isinstance(p, EkoStandPart) for p in parts):
            raise ValueError("All StandParts must be EkoStandPart objects!")
        self.parts = parts           # Swedish-style API
        self.Parts = self.parts      # Back-compat alias
        self.Site = site

        # Optional sanity: warn if Beech/Oak in non-southern regions
        if any(isinstance(p, EkoBeech) for p in self.parts) and self.Site.region not in ("Central", "South"):
            warnings.warn("Setting Beech stand outside of southern Sweden!")
        if any(isinstance(p, EkoOak) for p in self.parts) and self.Site.region not in ("Central", "South"):
            warnings.warn("Setting Oak stand outside of southern Sweden!")

        for p in self.parts:
            p.register_stand(self)

        # Initialize all derived metrics once
        self._assign_current_state_metrics()

    # ------------------------------------------------------------------
    # Competition / interaction metrics (HK etc.)
    # ------------------------------------------------------------------
    def _competition_metrics(self) -> None:
        """Compute per-part competition metrics from the current net state."""
        # Ensure own QMD is current
        for p in self.parts:
            p.QMD = self.getQMD(p.BA, p.stems)
        # Competition from all *other* parts
        for p in self.parts:
            BA_other = _safe_sum(q.BA for q in self.parts if q is not p)
            N_other = _safe_sum(q.stems for q in self.parts if q is not p)
            p.BAOtherSpecies = BA_other
            p.QMDOtherSpecies = self.getQMD(BA_other, N_other)
            denom = p.QMD if p.QMD > 0 else 1e-9
            # HK: diameter-based competition index, *not* height
            p.HK = (p.QMDOtherSpecies / denom) * BA_other

    # ------------------------------------------------------------------
    # Assign all current-state metrics: QMD, HK, VOL, stand totals
    # ------------------------------------------------------------------
    def _assign_current_state_metrics(self) -> None:
        """
        Recompute everything derived from the current net state:
          - per-part: QMD, competition metrics, volume (p.VOL)
          - stand totals: StandBA, StandStems, StandVOL
        """
        self._competition_metrics()
        for p in self.parts:
            p.VOL = p.getVolume(BA=p.BA, QMD=p.QMD, age=p.age,
                                stems=p.stems, HK=p.HK)

        self.StandBA = _safe_sum(p.BA for p in self.parts)
        self.StandStems = _safe_sum(p.stems for p in self.parts)
        self.StandVOL = _safe_sum(p.VOL for p in self.parts)

    # Back-compat alias some old code may call
    def _refresh_competition_vars(self) -> None:
        self._assign_current_state_metrics()

    # ------------------------------------------------------------------
    # Thinning (constant-QMD removal)
    # ------------------------------------------------------------------
    def thin(self, removals: dict):
        """
        removals: dict keyed by part.trädslag (enum) or species string; 
                  value = BA (m²/ha) to remove.
        Stems removed computed at current QMD (constant-QMD removal).
        """
        for p in self.parts:
            key_variants = (
                p.trädslag,
                getattr(p.trädslag, "value", None),
                str(getattr(p.trädslag, "value", "")),
            )
            BA_out = 0.0
            for k in key_variants:
                if k in removals:
                    BA_out = float(removals[k])
                    break

            BA_out = max(0.0, min(BA_out, p.BA))
            if BA_out > 0.0 and p.QMD > 0:
                # BA per tree = π * (d/200)²  (d in cm → m)
                stems_out = BA_out / (pi * (p.QMD / 200.0) ** 2)
                p.BA -= BA_out
                p.stems = max(0.0, p.stems - stems_out)

        self._assign_current_state_metrics()

    # ------------------------------------------------------------------
    # Five-year growth step: mortality + BAI (IGNORES HEIGHT TRAJECTORIES)
    # ------------------------------------------------------------------
    def grow(self, years: int = 5, apply_mortality: bool = True):
        """
        Advance the stand 'years' years (default 5).
        Mirrors the C++ updater (mortality + BAI applied in one step),
        but still ignores height trajectories (RK, RM2, etc.).
        Returns a dict keyed by species with N1, BA1, QMD1, VOL1.
        """
        # 0) Start-of-period metrics
        self._assign_current_state_metrics()

        # Snapshot starting metrics
        start_state = []
        for p in self.parts:
            start_state.append({
                "part": p,
                "BA": p.BA,
                "stems": p.stems,
                "age": p.age,
                "QMD": p.QMD,
                "VOL": p.getVolume(BA=p.BA, QMD=p.QMD, age=p.age,
                                   stems=p.stems, HK=p.HK),
            })
            p.VOL0 = start_state[-1]["VOL"]

        # 1) Mortality fractions + 2) Basal area increment (on start state)
        mortality = []
        for p in self.parts:
            if apply_mortality:
                (
                    BAQ_crowd,
                    _QMD_dead_crowd,
                    BAQ_other,
                    _QMD_dead_other,
                ) = p.getMortality(increment=years)
            else:
                BAQ_crowd = BAQ_other = 0.0

            p.getBAI5(
                ba_quotient_chronic_mortality=BAQ_crowd,
                ba_quotient_acute_mortality=BAQ_other,
            )
            mortality.append({
                "q_crowd": BAQ_crowd,
                "q_other": BAQ_other,
                "q_total": BAQ_crowd + BAQ_other,
            })

        # 3) Apply mortality + growth in one step (C++: ApplyMortalityAndGrowth)
        next_state = []
        for idx, p in enumerate(self.parts):
            q_total = mortality[idx]["q_total"] if apply_mortality else 0.0
            if apply_mortality:
                next_BA = (1.0 - q_total) * p.BA + p.BAI5
                next_stems = (1.0 - q_total) * p.stems
                next_age = p.age + years
            else:
                next_BA = p.BA
                next_stems = p.stems
                next_age = p.age

            next_QMD = self.getQMD(next_BA, next_stems)
            next_state.append({
                "part": p,
                "BA": next_BA,
                "stems": next_stems,
                "age": next_age,
                "QMD": next_QMD,
            })

        # 4) Competition & volumes on the post-mortality/post-growth state
        for idx, p in enumerate(self.parts):
            BA_other = _safe_sum(ns["BA"] for j, ns in enumerate(next_state) if j != idx)
            N_other = _safe_sum(ns["stems"] for j, ns in enumerate(next_state) if j != idx)
            QMD_other = self.getQMD(BA_other, N_other)
            HK_next = (QMD_other / (next_state[idx]["QMD"] if next_state[idx]["QMD"] > 0 else 1e-9)) * BA_other
            next_state[idx]["HK"] = HK_next

        for idx, p in enumerate(self.parts):
            ns = next_state[idx]
            ns["VOL"] = p.getVolume(
                BA=ns["BA"],
                QMD=ns["QMD"],
                age=ns["age"],
                stems=ns["stems"],
                HK=ns["HK"],
            )
            ns["volume_increment"] = ns["VOL"] - start_state[idx]["VOL"]

        # 5) Commit the new net state and return per-species summary
        period = {}
        for idx, p in enumerate(self.parts):
            ns = next_state[idx]
            p.gross_volume_increment = ns["volume_increment"]  # net increment after mortality
            p.volume_increment = ns["volume_increment"]

            p.BA = ns["BA"]
            p.stems = ns["stems"]
            p.QMD = ns["QMD"]
            p.age = ns["age"]
            p.HK = ns.get("HK", p.HK)
            p.VOL = ns["VOL"]

            key = p.trädslag.value   # e.g. "Tall", "Gran", "Björk", ...
            period.setdefault(key, []).append(
                {
                    "N1": p.stems,
                    "BA1": p.BA,
                    "QMD1": p.QMD,
                    "VOL1": p.VOL,
                    # Provenance (optional, handy for debugging)
                    "N0": start_state[idx]["stems"],
                    "BA0": start_state[idx]["BA"],
                    "QMD0": start_state[idx]["QMD"],
                    "VOL0": start_state[idx]["VOL"],
                }
            )

        # Finally refresh stand-level metrics for the new state
        self._assign_current_state_metrics()
        return period

    # Simple 5-year wrapper, matching the original model’s interface
    def grow5(self, mortality: bool = True):
        return self.grow(years=5, apply_mortality=mortality)
    
    def _volume_for(self, part, BA, QMD, age, stems, HK):
        BA   = part.BA   if BA   is None else BA
        age  = part.age  if age  is None else age
        stems = part.stems if stems is None else stems

        if QMD is None:
            QMD = self.getQMD(BA, stems)

        if HK is None:
            # Recompute competition index for this part from current stand state
            BA_other = sum(q.BA for q in self.Parts if q is not part)
            N_other  = sum(q.stems for q in self.Parts if q is not part)
            QMD_other = self.getQMD(BA_other, N_other)
            HK = (QMD_other / QMD) * BA_other if QMD > 0 else 0.0

        return part.getVolume(
            BA=BA,
            QMD=QMD,
            age=age,
            stems=stems,
            HK=HK,
        )

    

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
