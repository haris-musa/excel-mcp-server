"""CONVERT: the units Excel knows, with their metric and binary prefixes."""

from collections.abc import Callable
from dataclasses import dataclass

from excel_mcp.calc.registry import function
from excel_mcp.calc.values import NA, VALUE, FormulaError, Scalar, to_number


@dataclass(frozen=True)
class Unit:
    group: str
    factor: float  # in the base unit of its group
    prefix_power: int | None = None  # the power metric prefixes apply to; None: no prefixes


def _units(group: str, prefix_power: int | None, table: dict[str, float]) -> dict[str, Unit]:
    return {name: Unit(group, factor, prefix_power) for name, factor in table.items()}


def _fixed(group: str, table: dict[str, float]) -> dict[str, Unit]:
    return _units(group, None, table)


# Each factor is Excel's own: the size of the unit in the first unit named of its group.
UNITS: dict[str, Unit] = {
    **_units("mass", 1, {"g": 1.0, "u": 1.660538782e-24}),
    **_fixed(
        "mass",
        {
            "sg": 14593.902937206363,
            "lbm": 453.59237,
            "ozm": 28.349523125,
            "grain": 0.06479891,
            "cwt": 45359.237,
            "shweight": 45359.237,
            "uk_cwt": 50802.34544,
            "lcwt": 50802.34544,
            "hweight": 50802.34544,
            "stone": 6350.29318,
            "ton": 907184.74,
            "uk_ton": 1016046.9088,
            "LTON": 1016046.9088,
            "brton": 1016046.9088,
        },
    ),
    **_units("distance", 1, {"m": 1.0, "ang": 1e-10, "ly": 9460730472580800.0}),
    **_units("distance", 1, {"parsec": 3.0856775812815532e16, "pc": 3.0856775812815532e16}),
    **_fixed(
        "distance",
        {
            "mi": 1609.344,
            "Nmi": 1852.0,
            "in": 0.0254,
            "ft": 0.3048,
            "yd": 0.9144,
            "ell": 1.143,
            "Pica": 0.0003527777777777778,
            "Picapt": 0.0003527777777777778,
            "pica": 0.004233333333333334,
            "survey_mi": 1609.3472186944373,
        },
    ),
    **_units("time", 1, {"sec": 1.0, "s": 1.0}),
    **_fixed(
        "time",
        {"yr": 31557600.0, "day": 86400.0, "d": 86400.0, "hr": 3600.0, "mn": 60.0, "min": 60.0},
    ),
    **_units(
        "pressure", 1, {"Pa": 1.0, "p": 1.0, "atm": 101325.0, "at": 101325.0, "mmHg": 133.322}
    ),
    **_fixed("pressure", {"psi": 6894.757293168362, "Torr": 133.32236842105263}),
    **_units("force", 1, {"N": 1.0, "dyn": 1e-05, "dy": 1e-05, "pond": 0.00980665}),
    **_fixed("force", {"lbf": 4.4482216152605005}),
    **_units(
        "energy",
        1,
        {
            "J": 1.0,
            "e": 1e-07,
            "c": 4.184,
            "cal": 4.1868,
            "eV": 1.602176487e-19,
            "ev": 1.602176487e-19,
            "Wh": 3600.0,
            "wh": 3600.0,
        },
    ),
    **_fixed(
        "energy",
        {
            "HPh": 2684519.537696173,
            "hh": 2684519.537696173,
            "flb": 1.3558179483314003,
            "BTU": 1055.05585262,
            "btu": 1055.05585262,
        },
    ),
    **_units("power", 1, {"W": 1.0, "w": 1.0}),
    **_fixed("power", {"HP": 745.6998715822702, "h": 745.6998715822702, "PS": 735.49875}),
    **_units("magnetism", 1, {"T": 1.0, "ga": 0.0001}),
    **_units("volume", 1, {"l": 1.0, "L": 1.0, "lt": 1.0, "uk_pt": 0.56826125}),
    **_units("volume", 3, {"ang3": 1e-27, "ang^3": 1e-27, "m3": 1000.0, "m^3": 1000.0}),
    **_fixed(
        "volume",
        {
            "tsp": 0.00492892159375,
            "tspm": 0.005,
            "tbs": 0.01478676478125,
            "oz": 0.0295735295625,
            "cup": 0.2365882365,
            "pt": 0.473176473,
            "us_pt": 0.473176473,
            "qt": 0.946352946,
            "uk_qt": 1.1365225,
            "gal": 3.785411784,
            "uk_gal": 4.54609,
            "ft3": 28.316846592,
            "ft^3": 28.316846592,
            "in3": 0.016387064,
            "in^3": 0.016387064,
            "ly3": 8.467866646237152e50,
            "ly^3": 8.467866646237152e50,
            "mi3": 4168181825440.5796,
            "mi^3": 4168181825440.5796,
            "yd3": 764.554857984,
            "yd^3": 764.554857984,
            "Nmi3": 6352182208000.0,
            "Nmi^3": 6352182208000.0,
            "Pica3": 4.39039566186557e-08,
            "Picapt3": 4.39039566186557e-08,
            "Pica^3": 4.39039566186557e-08,
            "Picapt^3": 4.39039566186557e-08,
            "GRT": 2831.6846592,
            "regton": 2831.6846592,
            "MTON": 1132.67386368,
        },
    ),
    **_units("area", 2, {"ang2": 1e-20, "ang^2": 1e-20, "m2": 1.0, "m^2": 1.0}),
    **_units("area", 1, {"ar": 100.0}),
    **_fixed(
        "area",
        {
            "uk_acre": 4046.8564224,
            "us_acre": 4046.8726098742522,
            "ft2": 0.09290304,
            "ft^2": 0.09290304,
            "ha": 10000.0,
            "in2": 0.00064516,
            "in^2": 0.00064516,
            "ly2": 8.95054210748189e31,
            "ly^2": 8.95054210748189e31,
            "Morgen": 2500.0,
            "mi2": 2589988.110336,
            "mi^2": 2589988.110336,
            "Nmi2": 3429904.0,
            "Nmi^2": 3429904.0,
            "Pica2": 1.2445216049382715e-07,
            "Picapt2": 1.2445216049382715e-07,
            "Pica^2": 1.2445216049382715e-07,
            "Picapt^2": 1.2445216049382715e-07,
            "yd2": 0.83612736,
            "yd^2": 0.83612736,
        },
    ),
    **_units("information", 1, {"bit": 1.0, "byte": 8.0}),
    **_units("speed", 1, {"m/s": 1.0, "m/sec": 1.0, "m/h": 0.0002777777777777778}),
    **_units("speed", 1, {"m/hr": 0.0002777777777777778, "mph": 0.44704}),
    **_fixed("speed", {"admkn": 0.5147733333333333, "kn": 0.5144444444444445}),
}

PREFIXES = {
    "Y": 1e24,
    "Z": 1e21,
    "E": 1e18,
    "P": 1e15,
    "T": 1e12,
    "G": 1e9,
    "M": 1e6,
    "k": 1e3,
    "h": 1e2,
    "da": 1e1,
    "e": 1e1,
    "d": 1e-1,
    "c": 1e-2,
    "m": 1e-3,
    "u": 1e-6,
    "n": 1e-9,
    "p": 1e-12,
    "f": 1e-15,
    "a": 1e-18,
    "z": 1e-21,
    "y": 1e-24,
}
BINARY_PREFIXES = {
    "Yi": 2.0**80,
    "Zi": 2.0**70,
    "Ei": 2.0**60,
    "Pi": 2.0**50,
    "Ti": 2.0**40,
    "Gi": 2.0**30,
    "Mi": 2.0**20,
    "ki": 2.0**10,
}
# Temperatures go through kelvin: the function to it, and the one from it.
TEMPERATURES: dict[str, tuple[Callable[[float], float], Callable[[float], float]]] = {
    **dict.fromkeys(("C", "cel"), (lambda v: v + 273.15, lambda k: k - 273.15)),
    **dict.fromkeys(
        ("F", "fah"), (lambda v: (v - 32) / 1.8 + 273.15, lambda k: (k - 273.15) * 1.8 + 32)
    ),
    **dict.fromkeys(("K", "kel"), (lambda v: v, lambda k: k)),
    "Rank": (lambda v: v / 1.8, lambda k: k * 1.8),
    "Reau": (lambda v: v * 1.25 + 273.15, lambda k: (k - 273.15) * 0.8),
}
KELVIN = ("K", "kel")


@dataclass(frozen=True)
class Found:
    group: str
    factor: float
    prefix: float  # scales the value for a prefixed unit
    name: str


def _readings(name: str) -> list[Found]:
    """What a unit name can mean, plainest first: "pc" is the parsec or a picocalorie."""
    found: list[Found] = []
    if name in UNITS:
        found.append(Found(UNITS[name].group, UNITS[name].factor, 1.0, name))
    if name in TEMPERATURES:
        found.append(Found("temperature", 1.0, 1.0, name))
    for size in (1, 2):
        prefix, rest = name[:size], name[size:]
        scale = PREFIXES.get(prefix)
        if scale is None and rest in ("bit", "byte"):
            scale = BINARY_PREFIXES.get(prefix)
        if scale is None:
            continue
        unit = UNITS.get(rest)
        if unit is not None and unit.prefix_power is not None:
            found.append(Found(unit.group, unit.factor, scale**unit.prefix_power, rest))
        if rest in KELVIN and prefix in PREFIXES:
            found.append(Found("temperature", 1.0, scale, rest))
    return found


def _unit(value: Scalar) -> str:
    if isinstance(value, bool):
        raise FormulaError(VALUE)
    if not isinstance(value, str):
        raise FormulaError(NA)
    return value


@function("CONVERT", kind="scalar")
def convert(number: Scalar, from_unit: Scalar, to_unit: Scalar) -> float:
    if isinstance(number, bool):
        raise FormulaError(VALUE)
    value = to_number(number)
    names = _unit(from_unit), _unit(to_unit)
    pair = next(
        (
            (source, target)
            for source in _readings(names[0])
            for target in _readings(names[1])
            if source.group == target.group
        ),
        None,
    )
    if pair is None:
        raise FormulaError(NA)
    source, target = pair
    if source.group == "temperature":
        kelvin = TEMPERATURES[source.name][0](value * source.prefix)
        return TEMPERATURES[target.name][1](kelvin) / target.prefix
    return value * source.prefix * source.factor / (target.prefix * target.factor)
