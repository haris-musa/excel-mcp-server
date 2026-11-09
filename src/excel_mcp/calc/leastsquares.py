"""Least squares by orthogonalisation with pivoting, dropping columns that repeat others."""

import math
from dataclasses import dataclass

# A column whose part outside the chosen ones is this small beside its own length is dropped:
# found by bisecting recorded Excel results (the squared ratio is about 2**-41).
COLLINEAR = 2**-20.5


@dataclass(frozen=True)
class Fit:
    """A fit of y on columns. Dropped columns have ``kept`` False, coefficient 0 and no variance."""

    coefficients: list[float]
    kept: list[bool]
    residual: list[float]
    inverse_gram: list[list[float]]  # (X'X)^-1 over all columns, zero for the dropped ones


def least_squares(columns: list[list[float]], y: list[float]) -> Fit:
    count = len(columns)
    basis: list[list[float]] = []  # orthonormal vectors of the chosen columns
    chosen: list[int] = []
    triangle: list[list[float]] = []  # triangle[j][i]: chosen column j on basis vector i
    pending = list(range(count))
    while pending:
        parts = {i: _split(columns[i], basis) for i in pending}
        index = max(pending, key=lambda i: parts[i][1])
        pending.remove(index)
        components, new_length, rest = parts[index]
        if new_length > COLLINEAR * _norm(columns[index]):
            basis.append(_scaled(rest, 1 / new_length))
            chosen.append(index)
            triangle.append([*components, new_length])
    projections = [_dot(vector, y) for vector in basis]
    fitted = [sum(p * v[k] for p, v in zip(projections, basis, strict=True)) for k in range(len(y))]
    solved = _back_substitute(triangle, projections)
    inverse = _inverse(triangle)
    coefficients = [0.0] * count
    gram = [[0.0] * count for _ in range(count)]
    for a, column_a in enumerate(chosen):
        coefficients[column_a] = solved[a]
        for b, column_b in enumerate(chosen):
            gram[column_a][column_b] = sum(inverse[a][k] * inverse[b][k] for k in range(len(basis)))
    return Fit(
        coefficients,
        [i in chosen for i in range(count)],
        [a - b for a, b in zip(y, fitted, strict=True)],
        gram,
    )


def _split(column: list[float], basis: list[list[float]]) -> tuple[list[float], float, list[float]]:
    """The components of column on the basis, the length of what is left, and what is left."""
    rest = list(column)
    components = [0.0] * len(basis)
    for _ in range(2):
        for i, vector in enumerate(basis):
            shift = _dot(vector, rest)
            components[i] += shift
            rest = [r - shift * v for r, v in zip(rest, vector, strict=True)]
    return components, _norm(rest), rest


def _scaled(vector: list[float], factor: float) -> list[float]:
    return [v * factor for v in vector]


def _dot(first: list[float], second: list[float]) -> float:
    return math.fsum(a * b for a, b in zip(first, second, strict=True))


def _norm(vector: list[float]) -> float:
    return math.sqrt(_dot(vector, vector))


def _back_substitute(triangle: list[list[float]], target: list[float]) -> list[float]:
    """Solve R b = target, where R's column j is triangle[j] (nonzero up to row j)."""
    n = len(triangle)
    solution = [0.0] * n
    for row in range(n - 1, -1, -1):
        rest = sum(triangle[j][row] * solution[j] for j in range(row + 1, n))
        solution[row] = (target[row] - rest) / triangle[row][row]
    return solution


def _inverse(triangle: list[list[float]]) -> list[list[float]]:
    """R^-1 as rows, for the upper triangular R whose column j is triangle[j]."""
    n = len(triangle)
    columns = [
        _back_substitute(triangle, [1.0 if i == j else 0.0 for i in range(n)]) for j in range(n)
    ]
    return [[columns[j][i] for j in range(n)] for i in range(n)]
