"""Cell hyperlinks: places in the workbook, and plain http, https and mailto addresses."""

from typing import cast
from urllib.parse import urlsplit

from openpyxl.cell.cell import Cell
from openpyxl.utils import quote_sheetname
from openpyxl.workbook import Workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import split_top_level, unquote
from excel_mcp.inputs import InputModel
from excel_mcp.refs import parse_range

_SCHEMES = {"http", "https", "mailto"}
_MAX_TARGET = 2_079


class Link(InputModel):
    cell: str = Field(description="A cell of the written block, e.g. 'B2'.")
    target: str = Field(
        description="'https://...' or 'mailto:...', or a place in this workbook: "
        "\"#'Sheet 2'!A1\" or '#DefinedName'."
    )
    tooltip: str | None = Field(default=None, description="Text shown on hover.")


class LinkInfo(BaseModel):
    cell: str
    target: str
    tooltip: str | None = None


def set_link(cell: Cell, link: Link) -> None:
    workbook = cast(Workbook, cast(Worksheet, cell.parent).parent)
    target = link.target.strip()
    if len(target) > _MAX_TARGET or any(ord(char) < 32 for char in target):
        raise InvalidArgumentError(
            f"A link target must be one line of at most {_MAX_TARGET:,} characters."
        )
    if target.startswith("#"):
        place = _place(workbook, target[1:])
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=place, tooltip=link.tooltip)
    else:
        _check_address(target)
        cell.hyperlink = Hyperlink(ref=cell.coordinate, target=target, tooltip=link.tooltip)
    if not cell.has_style:
        cell.style = "Hyperlink"


def list_links(sheet: Worksheet) -> list[LinkInfo]:
    return [
        LinkInfo(
            cell=cell.coordinate,
            target=(cell.hyperlink.target or "")
            + ("#" + cell.hyperlink.location if cell.hyperlink.location else ""),
            tooltip=cell.hyperlink.tooltip,
        )
        for _, cell in sorted(sheet._cells.items())
        if cell.hyperlink
    ]


def _check_address(target: str) -> None:
    parts = urlsplit(target)
    has_destination = parts.path if parts.scheme == "mailto" else parts.netloc
    if parts.scheme not in _SCHEMES or not has_destination:
        raise InvalidArgumentError(
            f"Link {target!r} is not allowed: use a full http://, https:// or mailto: address, "
            "or '#Sheet!A1' for a place in this workbook. Other kinds, such as file: paths, "
            "network shares and javascript:, are refused."
        )


def _place(workbook: Workbook, place: str) -> str:
    parts = split_top_level(place, "!")
    if len(parts) == 1:
        if not _is_defined_name(workbook, place):
            raise InvalidArgumentError(
                f"Link target '#{place}' is neither 'Sheet!A1' nor a defined name. "
                "describe_workbook lists the defined names."
            )
        return place
    sheet = unquote(parts[0])
    if len(parts) != 2 or sheet not in workbook.sheetnames:
        raise InvalidArgumentError(
            f"Link sheet {parts[0]!r} not found. Available sheets: {workbook.sheetnames}."
        )
    return f"{quote_sheetname(sheet)}!{parse_range(parts[1])}"


def _is_defined_name(workbook: Workbook, name: str) -> bool:
    scopes = [workbook.defined_names, *(sheet.defined_names for sheet in workbook.worksheets)]
    return any(name in scope for scope in scopes)
