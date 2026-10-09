"""Removing the hyperlinks of an uploaded workbook that lead to files or network locations.

Such links are common in company workbooks and harmless until clicked, but a link to a
network share can hand the user's credentials to another host. They are removed, and what was
removed is reported, instead of refusing the upload. Links to web pages, mail addresses and
places in the workbook stay.
"""

import io
import re
import zipfile
from functools import partial
from typing import NamedTuple
from xml.sax.saxutils import unescape

from excel_mcp.errors import UnsafeFormulaError
from excel_mcp.links import is_allowed_address

_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship>)", re.DOTALL)
_ATTRIBUTE = re.compile(r"""([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")
_HREF = re.compile(r"""(\s)href\s*=\s*(?:"([^"]*)"|'([^']*)')""", re.IGNORECASE)
_LINK_ELEMENTS = ("hyperlink", "hlinkClick", "hlinkHover")
_SHOWN = 3
_MAX_SHOWN_TARGET = 80


class RemovedLink(NamedTuple):
    where: str
    target: str


def neutralise_links(content: bytes) -> tuple[bytes, list[RemovedLink]]:
    """The package without its unsafe external hyperlinks, and what was removed.

    Raises when a link cannot be removed without leaving a broken reference behind.
    """
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info) for info in infos}
    removed: list[RemovedLink] = []
    for name in [n for n in parts if n.casefold().endswith(".rels")]:
        removed += _remove_from_part(parts, name)
    for name in [n for n in parts if n.casefold().endswith(".vml")]:
        removed += _remove_from_vml(parts, name)
    if not removed:
        return content, []
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for info in infos:
            target.writestr(info.filename, parts[info.filename])
    return output.getvalue(), removed


def describe(removed: list[RemovedLink]) -> str:
    shown = ", ".join(f"{link.where} ({_shorten(link.target)})" for link in removed[:_SHOWN])
    more = f" and {len(removed) - _SHOWN} more" if len(removed) > _SHOWN else ""
    kind = "link" if len(removed) == 1 else "links"
    return f"Removed {len(removed)} {kind} to files or network locations: {shown}{more}."


def _shorten(target: str) -> str:
    return target if len(target) <= _MAX_SHOWN_TARGET else target[:_MAX_SHOWN_TARGET] + "..."


def _is_unsafe(relationship: str) -> tuple[str, str] | None:
    """The id and target of an external hyperlink relationship that must go."""
    values = {
        m[1]: unescape(m[2] if m[2] is not None else m[3])
        for m in _ATTRIBUTE.finditer(relationship)
    }
    external = values.get("TargetMode", "").casefold() == "external"
    if not external or not values.get("Type", "").casefold().endswith("/hyperlink"):
        return None
    target = values.get("Target", "")
    if target.startswith("#") or is_allowed_address(target):
        return None
    return values.get("Id", ""), target


def _owner(rels_name: str) -> str:
    directory, _, name = rels_name.rpartition("/")
    owner = name.removesuffix(".rels")
    parent = directory.removesuffix("_rels").removesuffix("/")
    return f"{parent}/{owner}" if parent else owner


def _remove_from_part(parts: dict[str, bytes], rels_name: str) -> list[RemovedLink]:
    try:
        text = parts[rels_name].decode("utf-8")
    except UnicodeDecodeError:
        return []
    doomed = {m[0]: found for m in _RELATIONSHIP.finditer(text) if (found := _is_unsafe(m[0]))}
    if not doomed:
        return []
    owner = _owner(rels_name)
    xml = parts[owner].decode("utf-8") if owner in parts else ""
    removed = []
    for relationship, (rel_id, target) in doomed.items():
        text = text.replace(relationship, "", 1)
        xml, where = _remove_references(xml, rel_id)
        removed.append(RemovedLink(where, target))
    parts[rels_name] = text.encode("utf-8")
    if owner in parts:
        parts[owner] = xml.encode("utf-8")
    return removed


def _remove_references(xml: str, rel_id: str) -> tuple[str, str]:
    """Drop the link elements that use a relationship id; returns the cell they were on."""
    where = "a shape or picture"
    quoted = re.escape(rel_id)
    for name in _LINK_ELEMENTS:
        pattern = re.compile(
            rf"""<((?:[\w.-]+:)?{name})\b(?=[^>]*\b(?:[\w.-]+:)?id\s*=\s*["']{quoted}["'])"""
            r"""[^>]*?(?:/>|>.*?</\1\s*>)""",
            re.DOTALL,
        )

        found: list[str] = []
        xml = pattern.sub(partial(_drop, found=found), xml)
        if found:
            where = found[0]
    if re.search(rf"""["']{quoted}["']""", xml):
        raise UnsafeFormulaError(
            "The workbook has a link to a file or network location that is used in a way "
            "that cannot be removed safely, so it cannot be uploaded."
        )
    return xml, where


def _drop(match: re.Match[str], *, found: list[str]) -> str:
    cell = re.search(r"""\bref\s*=\s*["']([^"']*)["']""", match[0])
    if cell:
        found.append(cell[1])
    return ""


def _remove_from_vml(parts: dict[str, bytes], name: str) -> list[RemovedLink]:
    text = parts[name].decode("utf-8", "ignore")
    removed: list[RemovedLink] = []

    def drop(match: re.Match[str]) -> str:
        target = unescape(match[2] if match[2] is not None else match[3])
        if target.startswith("#") or is_allowed_address(target):
            return match[0]
        removed.append(RemovedLink("a legacy drawing shape", target))
        return ""

    cleaned = _HREF.sub(drop, text)
    if removed:
        parts[name] = cleaned.encode("utf-8")
    return removed
