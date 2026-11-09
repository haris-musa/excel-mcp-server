"""Matrix functions: MMULT, MDETERM, MINVERSE, MUNIT."""

from excel_mcp.calc.functions.helpers import as_grid
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NUM,
    VALUE,
    FormulaError,
    Grid,
    Value,
    check_array_size,
    is_number,
    to_int,
)

Matrix = list[list[float]]


def _numbers(value: Value) -> Matrix:
    grid = as_grid(value)
    if not all(is_number(item) for item in grid.flat()):
        raise FormulaError(VALUE)
    return [[float(item) for item in row] for row in grid.rows]  # pyright: ignore[reportArgumentType]


def _square(value: Value) -> Matrix:
    matrix = _numbers(value)
    if len(matrix) != len(matrix[0]):
        raise FormulaError(VALUE)
    return matrix


@function("MMULT", array=True)
def mmult(first: Value, second: Value) -> Grid:
    left, right = _numbers(first), _numbers(second)
    if len(left[0]) != len(right):
        raise FormulaError(VALUE)
    check_array_size(len(left), len(right[0]))
    columns = list(zip(*right, strict=True))
    return Grid(
        [
            [naive_sum(a * b for a, b in zip(row, col, strict=True)) for col in columns]
            for row in left
        ]
    )


@function("MDETERM", array=True)
def mdeterm(array: Value) -> float:
    matrix = _square(array)
    size = len(matrix)
    determinant = 1.0
    for k in range(size):
        pivot = max(range(k, size), key=lambda r: abs(matrix[r][k]))
        if matrix[pivot][k] == 0:
            return 0.0
        if pivot != k:
            matrix[k], matrix[pivot] = matrix[pivot], matrix[k]
            determinant = -determinant
        determinant *= matrix[k][k]
        for r in range(k + 1, size):
            factor = matrix[r][k] / matrix[k][k]
            for c in range(k, size):
                matrix[r][c] -= factor * matrix[k][c]
    return determinant


@function("MINVERSE", array=True)
def minverse(array: Value) -> Grid:
    matrix = _square(array)
    size = len(matrix)
    check_array_size(size, size)
    work = [row + [1.0 if i == j else 0.0 for j in range(size)] for i, row in enumerate(matrix)]
    for k in range(size):
        pivot = max(range(k, size), key=lambda r: abs(work[r][k]))
        if work[pivot][k] == 0:
            raise FormulaError(NUM)
        work[k], work[pivot] = work[pivot], work[k]
        scale = work[k][k]
        work[k] = [item / scale for item in work[k]]
        for r in range(size):
            if r != k and work[r][k] != 0:
                factor = work[r][k]
                work[r] = [a - factor * b for a, b in zip(work[r], work[k], strict=True)]
    return Grid([row[size:] for row in work])


@function("MUNIT", kind="scalar")
def munit(dimension: Value) -> Grid:
    size = to_int(dimension)
    if size < 1:
        raise FormulaError(VALUE)
    check_array_size(size, size)
    return Grid([[1.0 if r == c else 0.0 for c in range(size)] for r in range(size)])
