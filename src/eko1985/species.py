"""Species-specific cohort implementations."""

from __future__ import annotations

from math import exp
from math import log as _math_log

from .base import EkoStandPart
from .enums import Trädslag


# -------------------------------------------------------------------
# Safe log wrapper
# -------------------------------------------------------------------
def log(x, eps: float = 1e-9) -> float:
    """
    Safe natural logarithm:
    - clamps non‑positive inputs to a small positive value (eps),
    - tolerates None and non‑numeric values by falling back to eps.

    This mimics the old C behavior where log(0) → very large negative
    (and thus exp(...) → ~0), but avoids Python's ValueError.
    """
    if x is None:
        return _math_log(eps)
    try:
        x = float(x)
    except (TypeError, ValueError):
        return _math_log(eps)
    if x <= 0.0:
        x = eps
    return _math_log(x)


class EkoSpruce(EkoStandPart):
    def __init__(self, ba, stems, age):
        super().__init__(ba, stems, age, Trädslag.GRAN)

    def getMortality(self, increment=5):
        """
        Bengtsson (1981) handwritten note (as in your code).
        Returns: BAQ due to crowding & its QMD; BAQ due to other & its QMD.
        """
        if self.stand is None:
            raise ValueError(
                "Mortality calculator requires EkoStand/EkoStandSite connected."
            )
        if self.stand.Site.region in ("North", "Central"):
            AKL = min((int(self.age) // 10) + 1, 17)
            crowding = (
                (-0.2748e-02 + 0.4493e-03 * self.BA + 0.2515e-04 * self.BA**2)
                * increment
                / 100.0
            )
            other = (
                -0.3150e-03 + 0.3337e-01 * AKL
            ) / 100.0  # fraction of BA over the period
        else:
            stems2 = min(self.stems, 2800)
            crowding = (
                (
                    0.1235e-01
                    + -0.2749e-02 * self.BA
                    + 0.8214e-04 * self.BA**2
                    + 0.2457e-04 * stems2
                    + -0.4498e-08 * stems2**2
                )
                * increment
                / 100.0
            )
            other = 0.36 / 100.0  # interpret as percent over the period

        if crowding > 1:
            crowding = 1.0
        elif crowding < 0:
            crowding = 0.0
        return crowding, 0.9 * self.QMD, other, self.QMD

    def getVolume(self, BA=None, QMD=None, age=None, stems=None, HK=None):
        if self.stand is None:
            raise ValueError(
                "Volume calculator cannot be called before part is connected to EkoStand/EkoStandSite."
            )
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10
        if self.stand.Site.region == "North":
            b1 = -0.065
            b2 = -2.05
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +0.362521e-02 * BA
                + 1.35682 * log(BA)
                - 1.47258 * QMD
                - 0.438770 * F4basal_area
                + 1.46910 * F4age
                - 0.314730 * log(stems)
                + 0.228700 * log(SIdm)
                + 0.118700e-01 * self.stand.Site.thinned
                + 0.254896e-02 * HK
                + 1.970094
            )
            return exp(lnVolume + 0.0388)
        elif self.stand.Site.region == "Central":
            b1 = -0.065
            b2 = -2.05
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +1.28359 * log(BA)
                - 0.380690 * F4basal_area
                + 1.21756 * F4age
                - 0.216690 * log(stems)
                + 0.350370 * log(SIdm)
                + 0.413000e-01 * self.stand.Site.HerbsGrassesNoFieldLayer
                + 0.362100e-01 * self.stand.Site.thinned
                + 0.268645e-02 * HK
                + 0.700490
            )
            return exp(lnVolume + 0.0563)
        else:
            b1 = -0.04
            b2 = -2.05
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +1.22886 * log(BA)
                - 0.349820 * F4basal_area
                + 0.485170 * F4age
                - 0.152050 * log(stems)
                + 0.337640 * log(SIdm)
                + 0.129800e-01 * self.stand.Site.thinned
                + 0.548055e-03 * HK
                + 0.584600
            )
            return exp(lnVolume + 0.0325)

    def getBAI5(
        self, ba_quotient_chronic_mortality=0.0, ba_quotient_acute_mortality=0.0
    ):
        if self.stand is None:
            raise ValueError("BAI calculator requires EkoStand/EkoStandSite connected.")
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10

        if self.stand.Site.region == "North":
            independent_vars = (
                -0.767477 * ba_quotient_chronic_mortality
                + -0.514297 * ba_quotient_acute_mortality
                + -1.43974 * self.QMD
                + -0.386338e-02 * self.HK
                + 0.204732 * self.stand.Site.fertilised
                + 0.186343 * self.stand.Site.vegcode
                + 0.392021e-01 * self.stand.Site.Bilberry_or_Cowberry
                + -0.807207e-01 * self.stand.Site.DrySoil
                + 0.833252
            )

            if SIdm < 160:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.736655e-02 * self.BA
                        + 0.875788 * log(self.BA)
                        - 0.642060e-04 * self.stems
                        + 0.125396 * log(self.stems)
                        + 0.159356e-02 * self.age
                        - 0.764340 * log(self.age)
                        - 0.594334e-02 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.187226e-01 * self.BA
                        + 0.855970 * log(self.BA)
                        + 0.106942e-03 * self.stems
                        + 0.107612 * log(self.stems)
                        + 0.321033e-02 * self.age
                        - 0.737062 * log(self.age)
                        - 0.206053e-01 * self.BAOtherSpecies
                    )
            elif SIdm < 200:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.191493e-01 * self.BA
                        + 0.942389 * log(self.BA)
                        - 0.145476e-03 * self.stems
                        + 0.158511 * log(self.stems)
                        + 0.289628e-02 * self.age
                        - 0.804217 * log(self.age)
                        - 0.125949e-01 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.255254e-01 * self.BA
                        + 0.955380 * log(self.BA)
                        - 0.642149e-04 * self.stems
                        + 0.164265 * log(self.stems)
                        + 0.554025e-02 * self.age
                        - 0.866520 * log(self.age)
                        - 0.889755e-02 * self.BAOtherSpecies
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.210737e-01 * self.BA
                        + 0.932275 * log(self.BA)
                        - 0.572335e-04 * self.stems
                        + 0.152017 * log(self.stems)
                        + 0.342622e-02 * self.age
                        - 0.811183 * log(self.age)
                        - 0.905176e-02 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.133941e-01 * self.BA
                        + 0.837783 * log(self.BA)
                        - 0.245946e-03 * self.stems
                        + 0.205142 * log(self.stems)
                        + 0.602419e-02 * self.age
                        - 0.862195 * log(self.age)
                        - 0.135941e-01 * self.BAOtherSpecies
                    )

            self.BAI5 = exp(dependent_vars + independent_vars + 0.0564)

        elif self.stand.Site.region == "Central":
            independent_vars = (
                -1.16597 * ba_quotient_chronic_mortality
                + -0.299327 * ba_quotient_acute_mortality
                + 0.783806e-01 * self.stand.Site.thinned_5y
                + 0.572131e-01 * self.stand.Site.vegcode
                + -0.112938e-01 * self.stand.Site.WetSoil
                + 0.546176e-01 * self.stand.Site.latitude
                + 0.332621e-01 * self.stand.Site.TAX77
            )

            if SIdm < 180:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.802837e-02 * self.BA
                        + 0.751220 * log(self.BA)
                        - 0.800241e-04 * self.stems
                        + 0.239814 * log(self.stems)
                        - 0.148757e-02 * self.age
                        - 0.476534 * log(self.age)
                        - 0.308451e-01 * self.BAOtherSpecies
                        - 4.02484
                    )
                else:
                    dependent_vars = (
                        -0.330623e-01 * self.BA
                        + 1.06539 * log(self.BA)
                        + 0.145290e-03 * self.stems
                        + 0.422450e-01 * log(self.stems)
                        + 0.110998e-01 * self.age
                        - 1.71468 * log(self.age)
                        - 0.236447e-01 * self.BAOtherSpecies
                        + 1.06383
                    )
            elif SIdm < 220:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.211171e-01 * self.BA
                        + 0.837241 * log(self.BA)
                        - 0.800241e-04 * self.stems
                        + 0.239814 * log(self.stems)
                        + 0.492578e-02 * self.age
                        - 0.839650 * log(self.age)
                        - 0.269523e-02 * self.BAOtherSpecies
                        - 2.91926
                    )
                else:
                    dependent_vars = (
                        -0.180419e-01 * self.BA
                        + 0.943986 * log(self.BA)
                        + 0.145290e-03 * self.stems
                        + 0.422450e-01 * log(self.stems)
                        + 0.525585e-02 * self.age
                        - 0.982261 * log(self.age)
                        - 0.786807e-02 * self.BAOtherSpecies
                        - 1.56544
                    )
            elif SIdm < 260:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.263745e-01 * self.BA
                        + 0.915196 * log(self.BA)
                        - 0.800241e-04 * self.stems
                        + 0.239814 * log(self.stems)
                        - 0.384471e-02 * self.age
                        - 0.847753 * log(self.age)
                        - 0.252559e-01 * self.BAOtherSpecies
                        + 2.85518
                    )
                else:
                    dependent_vars = (
                        -0.217674e-01 * self.BA
                        + 0.847682 * log(self.BA)
                        - 0.145290e-03 * self.stems
                        + 0.422450e-01 * log(self.stems)
                        + 0.101626e-01 * self.age
                        - 1.37782 * log(self.age)
                        - 0.268779e-01 * self.BAOtherSpecies
                        + 0.178428
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.244742e-01 * self.BA
                        + 0.787195 * log(self.BA)
                        - 0.800241e-04 * self.stems
                        + 0.239814 * log(self.stems)
                        + 0.371613e-02 * self.age
                        - 0.561641 * log(self.age)
                        - 0.298097e-01 * self.BAOtherSpecies
                        - 3.17570
                    )
                else:
                    dependent_vars = (
                        -0.239679e-01 * self.BA
                        + 0.924765 * log(self.BA)
                        + 0.145290e-03 * self.stems
                        + 0.422450e-01 * log(self.stems)
                        + 0.631561e-03 * self.age
                        - 0.893401 * log(self.age)
                        - 0.908286e-02 * self.BAOtherSpecies
                        - 1.46143
                    )

            self.BAI5 = exp(dependent_vars + independent_vars + 0.0712)

        else:
            independent_vars = (
                -0.780391 * ba_quotient_chronic_mortality
                + -0.252170 * ba_quotient_acute_mortality
                + -0.318464e-01 * self.stand.Site.thinned_5y
                + 0.778093e-01 * self.stand.Site.fertilised
                + 0.127135e-02 * SIdm
                + 0.262484e-01 * self.stand.Site.vegcode
                + -0.736690e-01 * self.stand.Site.DrySoil
                + -0.269193e-01 * self.stand.Site.latitude
                + -0.959785e-01 * self.stand.Site.TAX77
            )

            if SIdm < 220:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.149200e-01 * self.BA
                        + 0.794859 * log(self.BA)
                        - 0.120956e-03 * self.stems
                        + 0.255053 * log(self.stems)
                        - 0.720252 * log(self.age)
                        - 0.229139e-01 * self.BAOtherSpecies
                        + 1.52732
                    )
                else:
                    dependent_vars = (
                        -0.227763e-01 * self.BA
                        + 0.838105 * log(self.BA)
                        + 0.519813e-03 * self.stems
                        + 0.141232 * log(self.stems)
                        - 0.722723 * log(self.age)
                        - 0.237689e-01 * self.BAOtherSpecies
                        + 1.93218
                    )
            elif SIdm < 260:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.167127e-01 * self.BA
                        + 0.794738 * log(self.BA)
                        - 0.923244e-04 * self.stems
                        + 0.279717 * log(self.stems)
                        - 0.790588 * log(self.age)
                        - 0.187801e-01 * self.BAOtherSpecies
                        + 1.67230
                    )
                else:
                    dependent_vars = (
                        -0.167448e-01 * self.BA
                        + 0.835811 * log(self.BA)
                        - 0.995431e-04 * self.stems
                        + 0.258612 * log(self.stems)
                        - 0.931549 * log(self.age)
                        - 0.167010e-01 * self.BAOtherSpecies
                        + 2.34225
                    )
            elif SIdm < 300:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.221875e-01 * self.BA
                        + 0.832287 * log(self.BA)
                        - 0.110872e-03 * self.stems
                        + 0.271386 * log(self.stems)
                        - 0.735989 * log(self.age)
                        - 0.196143e-01 * self.BAOtherSpecies
                        + 1.50310
                    )
                else:
                    dependent_vars = (
                        -0.203970e-01 * self.BA
                        + 0.836890 * log(self.BA)
                        - 0.755155e-04 * self.stems
                        + 0.248563 * log(self.stems)
                        - 0.716504 * log(self.age)
                        - 0.151436e-01 * self.BAOtherSpecies
                        + 1.50719
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.243263e-01 * self.BA
                        + 0.902730 * log(self.BA)
                        - 0.706319e-04 * self.stems
                        + 0.198283 * log(self.stems)
                        - 0.713230 * log(self.age)
                        - 0.135840e-01 * self.BAOtherSpecies
                        + 1.71136
                    )
                else:
                    dependent_vars = (
                        -0.218319e-01 * self.BA
                        + 0.855200 * log(self.BA)
                        - 0.176554e-03 * self.stems
                        + 0.269091 * log(self.stems)
                        - 0.765104 * log(self.age)
                        - 0.180257e-01 * self.BAOtherSpecies
                        + 1.62508
                    )

            self.BAI5 = exp(dependent_vars + independent_vars + 0.0737)


class EkoPine(EkoStandPart):
    def __init__(self, ba, stems, age):
        super().__init__(ba, stems, age, Trädslag.TALL)

    def getMortality(self, increment=5):
        if self.stand is None:
            raise ValueError(
                "Mortality calculator requires EkoStand/EkoStandSite connected."
            )
        if self.stand.Site.region in ("North", "Central"):
            stems2 = min(self.stems, 2700)
            crowding = (
                (
                    0.3143e-01
                    + -0.6877e-02 * self.BA
                    + 0.2056e-03 * self.BA**2
                    + 0.2684e-04 * stems2
                    + -0.5092e-08 * stems2**2
                )
                * increment
                / 100.0
            )
            other = 0.35 / 100.0
        else:
            stems2 = min(self.stems, 4000)
            crowding = (
                (
                    -0.6766e-01
                    + -0.1283e-02 * self.BA
                    + 0.7748e-04 * self.BA**2
                    + 0.1441e-03 * stems2
                    + -0.1839e-07 * stems2**2
                )
                * increment
                / 100.0
            )
            other = 0.38 / 100.0

        if crowding > 1:
            crowding = 1.0
        elif crowding < 0:
            crowding = 0.0
        return crowding, 0.9 * self.QMD, other, self.QMD

    def getVolume(self, BA=None, QMD=None, age=None, stems=None, HK=None):
        if self.stand is None:
            raise ValueError(
                "Volume calculator cannot be called before part is connected to EkoStand/EkoStandSite."
            )
        SIdm = float(self.stand.Site.H100_Pine or 0.0) * 10

        if self.stand.Site.region == "North":
            b1 = -0.06
            b2 = -2.3
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +1.24296 * log(BA)
                - 0.472530 * F4basal_area
                + 1.05864 * F4age
                - 0.170140 * log(stems)
                + 0.247550 * log(SIdm)
                + 0.213800e-01 * self.stand.Site.thinned
                + 0.295300e-01 * self.stand.Site.thinned_5y
                + 0.510332e-02 * HK
                + 1.08339
            )
            return exp(lnVolume + 0.0275)

        elif self.stand.Site.region == "Central":
            b1 = -0.06
            b2 = -2.2
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +0.778157e-02 * BA
                + 1.14159 * log(BA)
                + 0.927460 * F4age
                - 0.166730 * log(stems)
                + 0.304900 * log(SIdm)
                + 0.270200e-01 * self.stand.Site.thinned
                + 0.292836e-02 * HK
                + 0.910330
            )
            return exp(lnVolume + 0.0273)

        else:
            b1 = -0.075
            b2 = -2.2
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +1.21272 * log(BA)
                - 0.299900 * F4basal_area
                + 1.01970 * F4age
                - 0.172300 * log(stems)
                + 0.369930 * log(SIdm)
                + 1.65136 * log(self.stand.Site.latitude)
                + 0.349200e-01 * log(self.stand.Site.altitude)
                - 0.197100e-01 * self.stand.Site.HerbsGrassesNoFieldLayer
                + 0.229100e-01 * self.stand.Site.thinned
                + 0.526017e-02 * HK
                - 6.46337
            )
            return exp(lnVolume + 0.0260)

    def getBAI5(
        self, ba_quotient_chronic_mortality=0.0, ba_quotient_acute_mortality=0.0
    ):
        if self.stand is None:
            raise ValueError("BAI calculator requires EkoStand/EkoStandSite connected.")
        SIdm = float(self.stand.Site.H100_Pine or 0.0) * 10

        if self.stand.Site.region == "North":
            independent_vars = (
                -0.598419 * ba_quotient_chronic_mortality
                + -0.486198 * ba_quotient_acute_mortality
                + -0.952624e-02 * self.HK
                + 0.674527e-01 * self.stand.Site.thinned_5y
                + 0.100135 * self.stand.Site.vegcode
                + -0.104076 * self.stand.Site.WetSoil
                + -0.329437e-01 * log(self.stand.Site.altitude)
                + 0.526479e-01 * self.stand.Site.TAX77
                + 0.164446
            )
            if SIdm < 160:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.342051e-01 * self.BA
                        + 0.757840 * log(self.BA)
                        - 0.161442e-03 * self.stems
                        + 0.367048 * log(self.stems)
                        + 0.313386e-02 * self.age
                        - 0.842335 * log(self.age)
                        - 0.157312e-01 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.222808e-01 * self.BA
                        + 0.707173 * log(self.BA)
                        - 0.407064e-03 * self.stems
                        + 0.386522 * log(self.stems)
                        + 0.309020e-02 * self.age
                        - 0.840856 * log(self.age)
                        - 0.168721e-01 * self.BAOtherSpecies
                    )
            elif SIdm < 200:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.264194e-01 * self.BA
                        + 0.759517 * log(self.BA)
                        - 0.172838e-03 * self.stems
                        + 0.354319 * log(self.stems)
                        + 0.282339e-02 * self.age
                        - 0.830969 * log(self.age)
                        - 0.920265e-02 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.215557e-01 * self.BA
                        + 0.678298 * log(self.BA)
                        - 0.223194e-03 * self.stems
                        + 0.345910 * log(self.stems)
                        + 0.230893e-02 * self.age
                        - 0.759426 * log(self.age)
                        - 0.129081e-01 * self.BAOtherSpecies
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.242773e-01 * self.BA
                        + 0.743286 * log(self.BA)
                        - 0.127080e-03 * self.stems
                        + 0.328240 * log(self.stems)
                        + 0.203892e-02 * self.age
                        - 0.756105 * log(self.age)
                        - 0.136312e-01 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.100435e-01 * self.BA
                        + 0.659451 * log(self.BA)
                        - 0.181913e-03 * self.stems
                        + 0.369130 * log(self.stems)
                        + 0.227817e-02 * self.age
                        - 0.793134 * log(self.age)
                        - 0.817145e-02 * self.BAOtherSpecies
                    )
            self.BAI5 = exp(dependent_vars + independent_vars + 0.0645)

        elif self.stand.Site.region == "Central":
            independent_vars = (
                -0.757422 * ba_quotient_chronic_mortality
                + -0.819721 * ba_quotient_acute_mortality
                + -0.156937e-01 * self.HK
                + 0.657419e-01 * self.stand.Site.fertilised
                + 0.208293e-02 * SIdm
                + 0.393424e-01 * self.stand.Site.vegcode
                + -0.787040e-01 * self.stand.Site.DrySoil
                + 0.952773e-01 * self.stand.Site.TAX77
                - 0.466279
            )
            if SIdm < 180:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.247769e-01 * self.BA
                        + 0.739123 * log(self.BA)
                        - 0.724080e-04 * self.stems
                        + 0.307962 * log(self.stems)
                        + 0.213813e-02 * self.age
                        - 0.730167 * log(self.age)
                        - 0.304936e-02 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.454216e-01 * self.BA
                        + 0.967594 * log(self.BA)
                        + 0.134748e-03 * self.stems
                        + 0.106405 * log(self.stems)
                        + 0.322181e-02 * self.age
                        - 0.559074 * log(self.age)
                        - 0.146382e-01 * self.BAOtherSpecies
                    )
            elif SIdm < 220:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.204976e-01 * self.BA
                        + 0.710569 * log(self.BA)
                        - 0.331436e-04 * self.stems
                        + 0.318007 * log(self.stems)
                        + 0.186999e-02 * self.age
                        - 0.732359 * log(self.age)
                        - 0.488064e-02 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        +0.144234e-01 * self.BA
                        + 0.304194 * log(self.BA)
                        - 0.111460e-02 * self.stems
                        + 0.628499 * log(self.stems)
                        + 0.545633e-02 * self.age
                        - 0.977317 * log(self.age)
                        - 0.126636e-01 * self.BAOtherSpecies
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.242132e-01 * self.BA
                        + 0.746931 * log(self.BA)
                        - 0.120517e-03 * self.stems
                        + 0.327216 * log(self.stems)
                        + 0.254795e-02 * self.age
                        - 0.758639 * log(self.age)
                        - 0.978754e-02 * self.BAOtherSpecies
                    )
                else:
                    dependent_vars = (
                        -0.126617e-01 * self.BA
                        + 0.599420 * log(self.BA)
                        - 0.405408e-03 * self.stems
                        + 0.472836 * log(self.stems)
                        + 0.455547e-02 * self.age
                        - 0.895734 * log(self.age)
                        - 0.106365e-01 * self.BAOtherSpecies
                    )
            self.BAI5 = exp(dependent_vars + independent_vars + 0.0507)

        else:
            independent_vars = (
                -1.04202 * ba_quotient_chronic_mortality
                + -0.637943 * ba_quotient_acute_mortality
                + -1.75160 * self.QMD
                + -0.592599e-02 * self.HK  # assuming HKD typo → HK
                + 0.637421e-01 * self.stand.Site.thinned_5y
                + 0.462966e-01 * self.stand.Site.fertilised
                + 0.522489e-01 * self.stand.Site.vegcode
                + -0.702839e-01 * self.stand.Site.DrySoil
                + -0.111568e-01 * self.stand.Site.latitude
                + -0.466973e-01 * self.stand.Site.TAX77
            )
            if SIdm < 160:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.497800e-01 * self.BA
                        + 1.19990 * log(self.BA)
                        + 0.114548e-04 * self.stems
                        + 0.164713 * log(self.stems)
                        - 0.884162e-03 * self.age
                        - 0.564604 * log(self.age)
                        - 0.153879e-01 * self.BAOtherSpecies
                        + 0.579562
                    )
                else:
                    dependent_vars = (
                        -0.302305e-01 * self.BA
                        + 0.938947 * log(self.BA)
                        + 0.563241e-03 * self.stems
                        + 0.148914 * log(self.stems)
                        + 0.419586e-02 * self.age
                        - 1.15586 * log(self.age)
                        - 0.138465e-01 * self.BAOtherSpecies
                        + 2.72773
                    )
            elif SIdm < 200:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.123212e-01 * self.BA
                        + 0.864851 * log(self.BA)
                        - 0.497769e-04 * self.stems
                        + 0.200066 * log(self.stems)
                        + 0.211976e-02 * self.age
                        - 0.821163 * log(self.age)
                        - 0.941390e-02 * self.BAOtherSpecies
                        + 1.59527
                    )
                else:
                    dependent_vars = (
                        -0.216126e-02 * self.BA
                        + 0.938131 * log(self.BA)
                        - 0.169034e-03 * self.stems
                        + 0.621225e-01 * log(self.stems)
                        + 0.305833e-02 * self.age
                        - 1.18279 * log(self.age)
                        - 0.439063e-03 * self.BAOtherSpecies
                        + 3.39954
                    )
            elif SIdm < 240:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.107718e-01 * self.BA
                        + 0.796896 * log(self.BA)
                        - 0.975686e-04 * self.stems
                        + 0.230066 * log(self.stems)
                        - 0.577520e-03 * self.age
                        - 0.570857 * log(self.age)
                        - 0.155230e-01 * self.BAOtherSpecies
                        + 0.784527
                    )
                else:
                    dependent_vars = (
                        -0.632941e-02 * self.BA
                        + 0.767710 * log(self.BA)
                        - 0.173551e-03 * self.stems
                        + 0.173044 * log(self.stems)
                        + 0.163026e-02 * self.age
                        - 0.945376 * log(self.age)
                        - 0.133437e-01 * self.BAOtherSpecies
                        + 2.49514
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.738511e-02 * self.BA
                        + 0.809028 * log(self.BA)
                        - 0.207393e-03 * self.stems
                        + 0.199179 * log(self.stems)
                        + 0.259619e-03 * self.age
                        - 0.663161 * log(self.age)
                        - 0.142082e-01 * self.BAOtherSpecies
                        + 1.27892
                    )
                else:
                    dependent_vars = (
                        -0.207497e-01 * self.BA
                        + 1.00931 * log(self.BA)
                        - 0.653755e-05 * self.stems
                        + 0.851371e-01 * log(self.stems)
                        - 0.307386e-02 * self.age
                        - 0.635182 * log(self.age)
                        - 0.110970e-01 * self.BAOtherSpecies
                        + 1.57124
                    )
            self.BAI5 = exp(dependent_vars + independent_vars + 0.0636)


class EkoBirch(EkoStandPart):
    def __init__(self, ba, stems, age):
        super().__init__(ba, stems, age, Trädslag.BJÖRK)

    def getMortality(self, increment=5):
        if self.stand is None:
            raise ValueError(
                "Mortality calculator requires EkoStand/EkoStandSite connected."
            )
        if self.stand.Site.region in ("North", "Central"):
            crowding = (-0.2513e-01 + 0.5489e-02 * self.BA) * increment / 100.0
            other = 0.78 / 100.0
        else:
            crowding = 0.04 * increment / 100.0
            other = 0.46 / 100.0

        if crowding > 1:
            crowding = 1.0
        elif crowding < 0:
            crowding = 0.0
        return crowding, 0.9 * self.QMD, other, self.QMD

    def getVolume(self, BA=None, QMD=None, age=None, stems=None, HK=None):
        if self.stand is None:
            raise ValueError(
                "Volume calculator cannot be called before part is connected to EkoStand/EkoStandSite."
            )
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10

        if self.stand.Site.region in ("North", "Central"):
            b1 = -0.035
            b2 = -2.05
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                +1.26244 * log(BA)
                - 0.459580 * F4basal_area
                + 0.540420 * F4age
                - 0.176040 * log(stems)
                + 0.201360 * log(SIdm)
                - 1.68251 * log(self.stand.Site.latitude)
                - 0.404000e-01 * log(self.stand.Site.altitude)
                + 0.757200e-01 * self.stand.Site.fertilised
                + 0.301200e-01 * self.stand.Site.thinned
                + 0.401844e-02 * HK
                + 8.44862
            )
            return exp(lnVolume + 0.0755)
        else:
            b1 = -0.07
            b2 = -2.1
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            lnVolume = (
                -0.786906e-02 * BA
                + 1.35254 * log(BA)
                - 1.30862 * QMD
                - 0.524630 * F4basal_area
                + 1.01779 * F4age
                - 0.254630 * log(stems)
                + 0.204880 * log(SIdm)
                + 2.75025 * log(self.stand.Site.latitude)
                + 0.774000e-01 * self.stand.Site.fertilised
                + 0.434800e-01 * self.stand.Site.thinned
                + 0.250449e-02 * HK
                - 9.38127
            )
            return exp(lnVolume + 0.0595)

    def getBAI5(
        self, ba_quotient_chronic_mortality=0.0, ba_quotient_acute_mortality=0.0
    ):
        if self.stand is None:
            raise ValueError("BAI calculator requires EkoStand/EkoStandSite connected.")
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10

        if self.stand.Site.region in ("North", "Central"):
            independent_vars = (
                -0.474848 * ba_quotient_chronic_mortality
                + -0.207333 * ba_quotient_acute_mortality
                + -0.202362e-02 * self.HK
                + 0.914442e-01 * self.stand.Site.thinned_5y
                + 0.176843 * self.stand.Site.fertilised
                + 0.256714 * self.stand.Site.vegcode
                + -0.488706e-01 * self.stand.Site.WetSoil
                + -0.139928e-01 * self.stand.Site.latitude
                + -0.462992 * self.stand.Site.altitude
                + 0.189383 * self.stand.Site.TAX77
            )
            if SIdm < 140:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        +0.281210e-02 * self.BA
                        + 0.718062 * log(self.BA)
                        - 0.264120e-03 * self.stems
                        + 0.360947 * log(self.stems)
                        - 0.513560 * log(self.age)
                        - 0.146581e-01 * self.BAOtherSpecies
                        - 0.768510
                    )
                else:
                    dependent_vars = (
                        +0.856585e-01 * self.BA
                        + 0.488507 * log(self.BA)
                        - 0.549010e-03 * self.stems
                        + 0.467588 * log(self.stems)
                        - 0.618645 * log(self.age)
                        - 0.477226e-02 * self.BAOtherSpecies
                        - 0.768510
                    )
            elif SIdm < 180:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        +0.831133e-02 * self.BA
                        + 0.660201 * log(self.BA)
                        - 0.161770e-03 * self.stems
                        + 0.361272 * log(self.stems)
                        - 0.609806 * log(self.age)
                        - 0.133204e-01 * self.BAOtherSpecies
                        - 0.355882
                    )
                else:
                    dependent_vars = (
                        +0.665931e-02 * self.BA
                        + 0.700295 * log(self.BA)
                        - 0.221485e-03 * self.stems
                        + 0.316196 * log(self.stems)
                        - 0.489888 * log(self.age)
                        - 0.246752e-01 * self.BAOtherSpecies
                        - 0.355882
                    )
            elif SIdm < 220:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.371203e-02 * self.BA
                        + 0.835899 * log(self.BA)
                        - 0.141238e-03 * self.stems
                        + 0.221611 * log(self.stems)
                        - 0.732659 * log(self.age)
                        - 0.131446e-01 * self.BAOtherSpecies
                        + 0.891049
                    )
                else:
                    dependent_vars = (
                        -0.134251e-02 * self.BA
                        + 0.838751 * log(self.BA)
                        - 0.237653e-03 * self.stems
                        + 0.192259 * log(self.stems)
                        - 0.707746 * log(self.age)
                        - 0.499067e-02 * self.BAOtherSpecies
                        + 0.891049
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.281602e-01 * self.BA
                        + 0.800357 * log(self.BA)
                        + 0.673284e-04 * self.stems
                        + 0.205233 * log(self.stems)
                        - 0.631139 * log(self.age)
                        - 0.176494e-01 * self.BAOtherSpecies
                        + 0.731245
                    )
                else:
                    dependent_vars = (
                        -0.177526e-01 * self.BA
                        + 0.814686 * log(self.BA)
                        + 0.781625e-04 * self.stems
                        + 0.183532 * log(self.stems)
                        - 0.593656 * log(self.age)
                        - 0.211444e-01 * self.BAOtherSpecies
                        + 0.731245
                    )
            self.BAI5 = exp(dependent_vars + independent_vars + 0.1642)

        else:
            independent_vars = (
                -0.617367 * ba_quotient_chronic_mortality
                + -0.350920 * ba_quotient_acute_mortality
                + -0.134245e-02 * self.HK
                + 0.277904 * self.stand.Site.fertilised
                + 0.154562 * self.stand.Site.vegcode
                + 0.554711e-01 * self.stand.Site.TAX77
            )
            if SIdm < 220:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.850224e-02 * self.BA
                        + 0.931518 * log(self.BA)
                        - 0.874696e-04 * self.stems
                        + 0.124964 * log(self.stems)
                        - 0.890226e-02 * self.age
                        - 0.498825 * log(self.age)
                        - 0.493910e-02 * self.BAOtherSpecies
                        - 0.135041
                    )
                else:
                    dependent_vars = (
                        +0.144427 * self.BA
                        + 0.332109 * log(self.BA)
                        - 0.457988e-03 * self.stems
                        + 0.474159 * log(self.stems)
                        + 0.922378e-02 * self.age
                        - 1.50315 * log(self.age)
                        - 0.116043e-01 * self.BAOtherSpecies
                        + 1.19213
                    )
            elif SIdm < 260:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        +0.129783e-01 * self.BA
                        + 0.688150 * log(self.BA)
                        - 0.158067e-03 * self.stems
                        + 0.304149 * log(self.stems)
                        + 0.411176e-02 * self.age
                        - 0.864501 * log(self.age)
                        - 0.533730e-02 * self.BAOtherSpecies
                        - 0.135041
                    )
                else:
                    dependent_vars = (
                        -0.235447e-01 * self.BA
                        + 0.962877 * log(self.BA)
                        + 0.103737e-03 * self.stems
                        + 0.186790 * log(self.stems)
                        - 0.127109e-02 * self.age
                        - 1.02854 * log(self.age)
                        - 0.849201e-02 * self.BAOtherSpecies
                        + 1.19213
                    )
            elif SIdm < 300:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.110984e-01 * self.BA
                        + 0.748193 * log(self.BA)
                        - 0.434390e-04 * self.stems
                        + 0.270476 * log(self.stems)
                        + 0.823613e-03 * self.age
                        - 0.718419 * log(self.age)
                        - 0.174522e-01 * self.BAOtherSpecies
                        - 0.135041
                    )
                else:
                    dependent_vars = (
                        -0.438786e-03 * self.BA
                        + 0.818427 * log(self.BA)
                        - 0.304146e-03 * self.stems
                        + 0.241055 * log(self.stems)
                        + 0.106700e-01 * self.age
                        - 1.16385 * log(self.age)
                        - 0.1978220e-01 * self.BAOtherSpecies
                        + 1.19213
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.204315e-01 * self.BA
                        + 0.792798 * log(self.BA)
                        - 0.179026e-03 * self.stems
                        + 0.316913 * log(self.stems)
                        + 0.262117e-02 * self.age
                        - 0.791796 * log(self.age)
                        - 0.146037e-01 * self.BAOtherSpecies
                        - 0.135041
                    )
                else:
                    dependent_vars = (
                        +0.255898e-02 * self.BA
                        + 0.730671 * log(self.BA)
                        + 0.256307e-04 * self.stems
                        + 0.256131 * log(self.stems)
                        + 0.126785e-01 * self.age
                        - 1.24005 * log(self.age)
                        - 0.341768e-02 * self.BAOtherSpecies
                        + 1.19213
                    )
            self.BAI5 = exp(dependent_vars + independent_vars + 0.1590)


class EkoBroadleaf(EkoStandPart):
    """Implementation for the grouped "other broadleaf" cohort."""

    def __init__(self, ba, stems, age):
        super().__init__(ba, stems, age, Trädslag.ÖV_LÖV)

    def getMortality(self, increment=5):
        if self.stand is None:
            raise ValueError(
                "Mortality calculator requires EkoStand/EkoStandSite connected."
            )
        if self.stand.Site.region in ("North", "Central"):
            crowding = (
                (-0.7277e-02 - 0.2456e-02 * self.BA + 0.1923e-03 * self.BA**2)
                * increment
                / 100.0
            )
            other = 0.5 / 100.0
        else:
            crowding = 0.04 * increment / 100.0
            other = 0.46 / 100.0

        if crowding > 1:
            crowding = 1.0
        elif crowding < 0:
            crowding = 0.0
        return crowding, 0.9 * self.QMD, other, self.QMD

    def getVolume(self, BA=None, QMD=None, age=None, stems=None, HK=None):
        if self.stand is None:
            raise ValueError(
                "Volume calculator cannot be called before part is connected to "
                "EkoStand/EkoStandSite."
            )
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10

        if self.stand.Site.region in ("North", "Central"):
            b1 = -0.04
            b2 = -2.3
            F4age = 1 - exp(b1 * age)
            F4basal_area = 1 - exp(b2 * BA)
            ln_volume = (
                1.26649 * log(BA)
                - 0.580030 * F4basal_area
                + 0.486310 * F4age
                - 0.172050 * log(stems)
                + 0.174930 * log(SIdm)
                - 1.51968 * log(self.stand.Site.latitude)
                - 0.368300e-01 * log(self.stand.Site.altitude)
                + 0.547400e-01 * self.stand.Site.thinned
                + 0.417126e-02 * HK
                + 7.79034
            )
            return exp(ln_volume + 0.0853)

        b1 = -0.075
        b2 = -2.1
        F4age = 1 - exp(b1 * age)
        F4basal_area = 1 - exp(b2 * BA)
        ln_volume = (
            -0.148700e-01 * BA
            + 1.29359 * log(BA)
            - 0.784820 * F4basal_area
            + 1.18741 * F4age
            - 0.135830 * log(stems)
            + 0.219890 * log(SIdm)
            + 2.02656 * log(self.stand.Site.latitude)
            + 0.242500e-01 * self.stand.Site.thinned
            + 0.859600e-01 * self.stand.StandBA
            + 0.509488e-03 * HK
            + 7.50102
        )
        return exp(ln_volume + 0.0671)

    def getBAI5(
        self, ba_quotient_chronic_mortality=0.0, ba_quotient_acute_mortality=0.0
    ):
        if self.stand is None:
            raise ValueError("BAI calculator requires EkoStand/EkoStandSite connected.")
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10

        if self.stand.Site.region in ("North", "Central"):
            independent_vars = (
                -0.345933 * ba_quotient_chronic_mortality
                - 0.138015 * self.stand.Site.vegcode
                - 0.650878e-01 * self.stand.Site.Bilberry_or_Cowberry
                - 0.175149e-01 * self.stand.Site.latitude
                - 0.570035e-03 * self.stand.Site.altitude
                + 0.151318 * self.stand.Site.TAX77
            )
            if SIdm < 160:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        +0.865166e-01 * self.BA
                        + 0.755603 * log(self.BA)
                        - 0.806548e-03 * self.stems
                        + 0.275974 * log(self.stems)
                        - 0.540881e-02 * self.age
                        - 0.117056 * log(self.age)
                        - 0.187866e-01 * self.BAOtherSpecies
                        - 1.18519
                    )
                else:
                    dependent_vars = (
                        +0.865166e-01 * self.BA
                        + 0.755603 * log(self.BA)
                        - 0.806548e-03 * self.stems
                        + 0.275974 * log(self.stems)
                        - 0.540881e-02 * self.age
                        - 0.117056 * log(self.age)
                        - 0.187866e-01 * self.BAOtherSpecies
                        - 0.952398
                    )
            elif SIdm < 200:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        -0.129773e-01 * self.BA
                        + 0.989525 * log(self.BA)
                        - 0.715363e-04 * self.stems
                        + 0.490676e-01 * log(self.stems)
                        + 0.218728e-02 * self.age
                        - 0.944317 * log(self.age)
                        - 0.143834e-01 * self.BAOtherSpecies
                        + 2.78296
                    )
                else:
                    dependent_vars = (
                        -0.129773e-01 * self.BA
                        + 0.989525 * log(self.BA)
                        - 0.715363e-04 * self.stems
                        + 0.490676e-01 * log(self.stems)
                        + 0.218728e-02 * self.age
                        - 0.944317 * log(self.age)
                        - 0.143834e-01 * self.BAOtherSpecies
                        + 2.87671
                    )
            elif SIdm < 240:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        +0.517826e-01 * self.BA
                        + 0.768565 * log(self.BA)
                        - 0.381320e-03 * self.stems
                        + 0.201267 * log(self.stems)
                        + 0.131078e-02 * self.age
                        - 0.831523 * log(self.age)
                        - 0.122796e-01 * self.BAOtherSpecies
                        + 1.65650
                    )
                else:
                    dependent_vars = (
                        +0.517826e-01 * self.BA
                        + 0.768565 * log(self.BA)
                        - 0.381320e-03 * self.stems
                        + 0.201267 * log(self.stems)
                        + 0.131078e-02 * self.age
                        - 0.831523 * log(self.age)
                        - 0.122796e-01 * self.BAOtherSpecies
                        + 1.59209
                    )
            else:
                if not self.stand.Site.thinned:
                    dependent_vars = (
                        +0.243920e-02 * self.BA
                        + 0.857832 * log(self.BA)
                        - 0.949555e-04 * self.stems
                        + 0.192173 * log(self.stems)
                        - 0.292753e-02 * self.age
                        - 0.570009 * log(self.age)
                        - 0.240816e-01 * self.BAOtherSpecies
                        + 0.916942
                    )
                else:
                    dependent_vars = (
                        +0.243920e-02 * self.BA
                        + 0.857832 * log(self.BA)
                        - 0.949555e-04 * self.stems
                        + 0.192173 * log(self.stems)
                        - 0.292753e-02 * self.age
                        - 0.570009 * log(self.age)
                        - 0.240816e-01 * self.BAOtherSpecies
                        + 1.17865
                    )
            self.BAI5 = exp(dependent_vars + independent_vars + 0.1648)
            return

        independent_vars = (
            -1.20049 * ba_quotient_chronic_mortality
            - 0.367064 * ba_quotient_acute_mortality
            + 0.125048 * self.stand.Site.thinned_5y
            + 0.246684 * self.stand.Site.fertilised
            + 0.141955 * self.stand.Site.vegcode
            + 0.354866e-01 * self.stand.Site.latitude
            - 0.361988e-03 * self.stand.Site.altitude
        )
        if SIdm < 240:
            if not self.stand.Site.thinned:
                dependent_vars = (
                    +0.857153 * log(self.BA)
                    - 0.541853e-04 * self.stems
                    + 0.152684 * log(self.stems)
                    - 0.803085e-02 * self.age
                    - 0.570230 * log(self.age)
                    - 0.100518 * log(self.BAOtherSpecies)
                    - 1.93895
                )
            else:
                dependent_vars = (
                    +0.857153 * log(self.BA)
                    - 0.541853e-04 * self.stems
                    + 0.152684 * log(self.stems)
                    - 0.803085e-02 * self.age
                    - 0.570230 * log(self.age)
                    - 0.100518 * log(self.BAOtherSpecies)
                    - 2.01960
                )
        elif SIdm < 280:
            if not self.stand.Site.thinned:
                dependent_vars = (
                    +0.794405 * log(self.BA)
                    - 0.247009 * self.stems
                    + 0.202344 * log(self.stems)
                    - 0.250423 * self.age
                    - 0.669629 * log(self.age)
                    - 0.101205 * log(self.BAOtherSpecies)
                    - 1.93895
                )
            else:
                dependent_vars = (
                    +0.794405 * log(self.BA)
                    - 0.247009 * self.stems
                    + 0.202344 * log(self.stems)
                    - 0.250423 * self.age
                    - 0.669629 * log(self.age)
                    - 0.101205 * log(self.BAOtherSpecies)
                    - 2.01960
                )
        elif SIdm < 320:
            if not self.stand.Site.thinned:
                dependent_vars = (
                    +0.782374 * log(self.BA)
                    - 0.125111e-03 * self.stems
                    + 0.239626 * log(self.stems)
                    - 0.787146e-03 * self.age
                    - 0.733575 * log(self.age)
                    - 0.823802e-01 * log(self.BAOtherSpecies)
                    - 1.93895
                )
            else:
                dependent_vars = (
                    +0.782374 * log(self.BA)
                    - 0.125111e-03 * self.stems
                    + 0.239626 * log(self.stems)
                    - 0.787146e-03 * self.age
                    - 0.733575 * log(self.age)
                    - 0.823802e-01 * log(self.BAOtherSpecies)
                    - 2.01960
                )
        else:
            if not self.stand.Site.thinned:
                dependent_vars = (
                    +0.771398 * log(self.BA)
                    + 0.427071e-04 * self.stems
                    + 0.167037 * log(self.stems)
                    - 0.190695e-02 * self.age
                    - 0.587696 * log(self.age)
                    - 0.113489 * log(self.BAOtherSpecies)
                    - 1.93895
                )
            else:
                dependent_vars = (
                    +0.771398 * log(self.BA)
                    + 0.427071e-04 * self.stems
                    + 0.167037 * log(self.stems)
                    - 0.190695e-02 * self.age
                    - 0.587696 * log(self.age)
                    - 0.113489 * log(self.BAOtherSpecies)
                    - 2.01960
                )
        self.BAI5 = exp(dependent_vars + independent_vars + 0.1734)


class EkoBeech(EkoStandPart):
    def __init__(self, ba, stems, age):
        super().__init__(ba, stems, age, Trädslag.BOK)

    def getMortality(self, increment=5):
        if self.stand is None:
            raise ValueError(
                "Mortality calculator requires EkoStand/EkoStandSite connected."
            )
        if self.stand.Site.region in ("North", "Central"):
            crowding = (
                (-0.7277e-02 + -0.2456e-02 * self.BA + 0.1923e-03 * self.BA**2)
                * increment
                / 100.0
            )
            other = 0.5 / 100.0
        else:
            crowding = 0.04 * increment / 100.0
            other = 0.46 / 100.0

        if crowding > 1:
            crowding = 1.0
        elif crowding < 0:
            crowding = 0.0
        return crowding, 0.9 * self.QMD, other, self.QMD

    def getVolume(self, BA=None, QMD=None, age=None, stems=None, HK=None):
        if self.stand is None:
            raise ValueError(
                "Volume calculator cannot be called before part is connected to EkoStand/EkoStandSite."
            )
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10
        b1 = -0.02
        b2 = -2.3
        F4age = 1 - exp(b1 * age)
        F4basal_area = 1 - exp(b2 * BA)
        lnVolume = (
            -0.111600e-01 * BA
            + 1.30527 * log(BA)
            - 0.676190 * F4basal_area
            + 0.490740 * F4age
            - 0.151930 * log(stems)
            - 0.572600e-01 * log(SIdm)
            + 0.628000e-01 * self.stand.Site.thinned
            + 0.203927e-02 * HK
            + 2.85509
        )
        return exp(lnVolume + 0.0392)

    def getBAI5(
        self, ba_quotient_chronic_mortality=0.0, ba_quotient_acute_mortality=0.0
    ):
        if self.stand is None:
            raise ValueError("BAI calculator requires EkoStand/EkoStandSite connected.")
        SIdm = self.stand.Site.H100_Spruce * 10

        independent_vars = (
            -0.862301 * ba_quotient_acute_mortality + 0.162579e-02 * SIdm + 0.538943
        )

        if SIdm < 310:
            if not self.stand.Site.thinned:
                dependent_vars = (
                    +0.948126 * log(self.BA)
                    + 0.563620e-01 * log(self.stems)
                    - 0.751665 * log(self.age)
                    - 0.163302e-01 * self.BAOtherSpecies
                )
            else:
                dependent_vars = (
                    +0.948126 * log(self.BA)
                    + 0.563620e-01 * log(self.stems)
                    - 0.751665 * log(self.age)
                    - 0.163302e-01 * self.BAOtherSpecies
                    + 0.887110e-01
                )
        else:
            if not self.stand.Site.thinned:
                dependent_vars = (
                    +0.821914 * log(self.BA)
                    + 0.102770 * log(self.stems)
                    - 0.753735 * log(self.age)
                    - 0.163641e-01 * self.BAOtherSpecies
                )
            else:
                dependent_vars = (
                    +0.821914 * log(self.BA)
                    + 0.102770 * log(self.stems)
                    - 0.753735 * log(self.age)
                    - 0.163641e-01 * self.BAOtherSpecies
                    + 0.887110e-01
                )

        self.BAI5 = exp(dependent_vars + independent_vars + 0.1379)


class EkoOak(EkoStandPart):
    def __init__(self, ba, stems, age):
        super().__init__(ba, stems, age, Trädslag.EK)

    def getMortality(self, increment=5):
        if self.stand is None:
            raise ValueError(
                "Mortality calculator requires EkoStand/EkoStandSite connected."
            )
        if self.stand.Site.region in ("North", "Central"):
            crowding = (
                (-0.7277e-02 + -0.2456e-02 * self.BA + 0.1923e-03 * self.BA**2)
                * increment
                / 100.0
            )
            other = 0.5 / 100.0
        else:
            crowding = 0.04 * increment / 100.0
            other = 0.46 / 100.0

        if crowding > 1:
            crowding = 1.0
        elif crowding < 0:
            crowding = 0.0
        return crowding, 0.9 * self.QMD, other, self.QMD

    def getVolume(self, BA=None, QMD=None, age=None, stems=None, HK=None):
        if self.stand is None:
            raise ValueError(
                "Volume calculator cannot be called before part is connected to EkoStand/EkoStandSite."
            )
        SIdm = float(self.stand.Site.H100_Spruce or 0.0) * 10
        b1 = -0.055
        b2 = -2.3
        F4age = 1 - exp(b1 * age)
        F4basal_area = 1 - exp(b2 * BA)
        lnVolume = (
            -0.106300e-01 * BA
            + 1.27353 * log(BA)
            - 0.463790 * F4basal_area
            + 0.801580 * F4age
            - 0.157080 * log(stems)
            + 0.159030 * log(SIdm)
            + 0.503200e-01 * self.stand.Site.thinned
            + 0.188030e-02 * HK
            + 1.40608
        )
        return exp(lnVolume + 0.0756)

    def getBAI5(
        self, ba_quotient_chronic_mortality=0.0, ba_quotient_acute_mortality=0.0
    ):
        if self.stand is None:
            raise ValueError("BAI calculator requires EkoStand/EkoStandSite connected.")
        SIdm = self.stand.Site.H100_Spruce * 10

        independent_vars = -0.389169 * ba_quotient_acute_mortality - 0.609667
        if SIdm < 280:
            dependent_vars = (
                +0.896599 * log(self.BA)
                + 0.199354 * log(self.stems)
                - 0.842665 * log(self.age)
                - 0.146432e-01 * self.BAOtherSpecies
            )
        elif SIdm < 320:
            dependent_vars = (
                +0.847420 * log(self.BA)
                + 0.144495 * log(self.stems)
                - 0.727278 * log(self.age)
                - 0.222990e-01 * self.BAOtherSpecies
            )
        else:
            dependent_vars = (
                +0.851362 * log(self.BA)
                + 0.128100 * log(self.stems)
                - 0.667346 * log(self.age)
                - 0.199705e-01 * self.BAOtherSpecies
            )
        self.BAI5 = exp(dependent_vars + independent_vars + 0.1618)


__all__ = [
    "EkoSpruce",
    "EkoPine",
    "EkoBirch",
    "EkoBeech",
    "EkoOak",
    "EkoBroadleaf",
]
