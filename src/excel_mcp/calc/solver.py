"""Goal Seek: the input that makes a function reach a value, with Excel's stopping rules."""

import math
from collections.abc import Callable
from dataclasses import dataclass

MAX_ITERATIONS = 100  # Excel's default for iterative calculation
MAX_CHANGE = 0.001  # ... and the largest distance from the goal that counts as reached
_HALVINGS = 8


@dataclass
class Point:
    x: float
    value: float  # what the function returned at x
    miss: float  # value minus the goal


@dataclass
class Outcome:
    best: Point
    iterations: int

    @property
    def reached(self) -> bool:
        return abs(self.best.miss) < MAX_CHANGE


def seek(evaluate: Callable[[float], float | None], start: float, goal: float) -> Outcome | None:
    """Look for x where ``evaluate(x)`` is within `MAX_CHANGE` of ``goal``, starting at ``start``.

    ``evaluate`` returns None where the function has no number (an error value). The search
    takes secant steps until the goal is between two values, then narrows that interval
    (Illinois method). None when the function has no number at ``start``.
    """
    iterations = 0
    best: Point | None = None

    def probe(x: float) -> Point | None:
        nonlocal iterations, best
        iterations += 1
        value = evaluate(x)
        if value is None or not math.isfinite(value):
            return None
        point = Point(x, value, value - goal)
        if best is None or abs(point.miss) < abs(best.miss):
            best = point
        return point

    def probe_towards(origin: Point, x: float) -> Point | None:
        """Probe x; where the function has no number, try half way back towards ``origin``."""
        for _ in range(_HALVINGS):
            if (point := probe(x)) is not None or iterations >= MAX_ITERATIONS:
                return point
            x = (x + origin.x) / 2
        return None

    older = probe(start)
    if older is None:
        return None
    newer = probe_towards(older, start + (0.01 * abs(start) or 0.01))
    ends: list[list[float]] | None = None  # [x, scaled miss] of two points around the goal
    side = 0
    while newer is not None and iterations < MAX_ITERATIONS and abs(newer.miss) >= MAX_CHANGE:
        x = _secant(older, newer) if ends is None else _interpolate(ends)
        if x == newer.x or (ends is not None and abs(ends[0][0] - ends[1][0]) < 1e-15):
            break
        point = probe_towards(newer, x)
        if point is None:
            break
        if ends is None:
            if math.copysign(1, point.miss) != math.copysign(1, newer.miss):
                ends = [[newer.x, newer.miss], [point.x, point.miss]]
                side = 1
        else:
            side = _narrow(ends, point, side)
        older, newer = newer, point
    assert best is not None
    return Outcome(best, iterations)


def _secant(older: Point, newer: Point) -> float:
    """Where the line through two points reaches the goal; a flat line goes on in its direction."""
    slope = (newer.miss - older.miss) / (newer.x - older.x)
    if slope == 0 or not math.isfinite(slope):
        return newer.x + 2 * (newer.x - older.x)
    return newer.x - newer.miss / slope


def _interpolate(ends: list[list[float]]) -> float:
    """Where the line between the ends reaches the goal; the middle when that is hardly
    different from an end, as when one miss is astronomically larger than the other."""
    (x0, m0), (x1, m1) = ends
    x = x1 - m1 * (x1 - x0) / (m1 - m0)
    if min(abs(x - x0), abs(x - x1)) < 1e-9 * abs(x1 - x0):
        return (x0 + x1) / 2
    return x


def _narrow(ends: list[list[float]], point: Point, side: int) -> int:
    """Replace the end of the interval on the same side of the goal as ``point``. An end kept
    twice in a row has its miss halved, so that the interval shrinks from both sides."""
    keep, drop = (0, 1) if math.copysign(1, point.miss) == math.copysign(1, ends[1][1]) else (1, 0)
    kept_side = 1 if drop == 1 else -1
    if side == kept_side:
        ends[keep][1] /= 2
    ends[drop] = [point.x, point.miss]
    return kept_side
