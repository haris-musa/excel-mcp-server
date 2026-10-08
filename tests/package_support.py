"""Helpers for tests that look inside the zip of a workbook the server saved."""

import re
import shutil
import zipfile
from pathlib import Path
from posixpath import dirname, join, normpath

FIXTURES = Path(__file__).parent / "fixtures"
RELATIONSHIPS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def copy_fixture(files: Path, fixture: str, name: str = "book.xlsx") -> Path:
    return Path(shutil.copy(FIXTURES / fixture, files / name))


def read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def text(parts: dict[str, bytes], name: str) -> str:
    return parts[name].decode("utf-8")


def relationships(parts: dict[str, bytes], source: str) -> dict[str, tuple[str, str]]:
    """The relationships of a part as {id: (type, target part name)}."""
    directory, _, name = source.rpartition("/")
    rels_name = f"{directory}/_rels/{name}.rels" if directory else f"_rels/{name}.rels"
    if rels_name not in parts:
        return {}
    found = {}
    for tag in re.findall(r"<Relationship\b[^>]*>", text(parts, rels_name)):
        values = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        target = values["Target"]
        if values.get("TargetMode") != "External":
            target = normpath(
                target[1:] if target.startswith("/") else join(dirname(source), target)
            )
        found[values["Id"]] = (values["Type"].rsplit("/", 1)[-1], target)
    return found


def sheet_part(parts: dict[str, bytes], sheet: str) -> str:
    workbook = text(parts, "xl/workbook.xml")
    for tag in re.findall(r"<sheet\b[^>]*>", workbook):
        values = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        if values["name"] == sheet:
            return relationships(parts, "xl/workbook.xml")[values["r:id"]][1]
    raise KeyError(sheet)


def sheet_ids(parts: dict[str, bytes]) -> dict[str, int]:
    workbook = text(parts, "xl/workbook.xml")
    return {
        values["name"]: int(values["sheetId"])
        for values in (
            dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
            for tag in re.findall(r"<sheet\b[^>]*>", workbook)
        )
    }


def content_type(parts: dict[str, bytes], name: str) -> str | None:
    types = text(parts, "[Content_Types].xml")
    override = re.search(rf'PartName="/{re.escape(name)}" ContentType="([^"]*)"', types)
    if override:
        return override[1]
    extension = name.rpartition(".")[2]
    default = re.search(rf'Extension="{extension}" ContentType="([^"]*)"', types)
    return default[1] if default else None


def assert_package_is_consistent(parts: dict[str, bytes]) -> None:
    """Every part has a content type, and every relationship points at a part that exists."""
    for name in parts:
        if not name.endswith(".rels"):
            assert content_type(parts, name), f"{name} has no content type"
    for name in parts:
        if name.endswith(".rels"):
            directory, _, rels_file = (
                name.rpartition("/_rels/") if "/_rels/" in name else ("", "", name[len("_rels/") :])
            )
            source = (
                f"{directory}/{rels_file[: -len('.rels')]}"
                if directory
                else rels_file[: -len(".rels")]
            )
            for rel_id, (_, target) in relationships(parts, source).items():
                assert target in parts or "://" in target, (name, rel_id, target)
