"""Workbooks, damaged copies of them and schema-driven random arguments for the fuzz test."""

import base64
import random
import re
import zipfile
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

FIXTURES = Path(__file__).parent / "fixtures"
SEEDS = [
    "excel_chartex.xlsx",
    "excel_controls.xlsx",
    "excel_dynamic_arrays.xlsx",
    "excel_objects.xlsx",
    "excel_pivot_groups.xlsx",
    "excel_pivot_source.xlsx",
    "excel_shapes.xlsx",
    "excel_slicers.xlsx",
    "excel_sparklines.xlsx",
    "macro01.xlsm",
]
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg=="
)


def rich_workbook(path: Path) -> None:
    """Data, formulas, a table, merges, notes, links, names and sheets of every state."""
    workbook = Workbook()
    data = workbook.worksheets[0]
    data.title = "Data"
    data.append(["Region", "Product", "Units", "Price", "Date"])
    for index in range(1, 41):
        data.append(
            [
                ["North", "South", "East"][index % 3],
                f"P{index % 5}",
                index,
                index * 1.5,
                f"2024-0{index % 9 + 1}-01",
            ]
        )
    data["G1"], data["G2"], data["G3"] = "=SUM(C2:C41)", "=G1*2", '=INDIRECT("G"&1)'
    data["G4"], data["G5"] = "=1/0", '=TEXT(G1,"#,##0.00")'
    data["H1"] = "=A1&H2"
    data["H2"] = "=H1"
    data.merge_cells("J1:K2")
    data["J5"].comment = Comment("a note", "me")
    data["J6"].hyperlink = "https://example.com"
    data.add_table(Table(displayName="Sales", ref="A1:E41"))
    validation = DataValidation(type="list", formula1='"a,b"')
    data.add_data_validation(validation)
    validation.add("L1")
    data.freeze_panes = "B2"
    hidden = workbook.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    hidden["A1"] = 1
    very = workbook.create_sheet("Very Hidden")
    very.sheet_state = "veryHidden"
    locked = workbook.create_sheet("Locked")
    locked["A1"] = "x"
    locked.protection.sheet = True
    workbook.create_sheet("Empty")
    workbook.defined_names["Rate"] = DefinedName("Rate", attr_text="Data!$D$2")
    workbook.save(path)


def large_merge_workbook(path: Path) -> None:
    """One merge over the whole sheet; openpyxl would make a cell object for each of its cells."""
    workbook = Workbook()
    workbook.worksheets[0]["A1"] = 1
    plain = path.with_suffix(".plain")
    workbook.save(plain)
    with zipfile.ZipFile(plain) as package:
        sheet = package.read("xl/worksheets/sheet1.xml")
    merged = sheet.replace(
        b"</sheetData>",
        b'</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>',
    )
    _rewrite(plain, path, {"xl/worksheets/sheet1.xml": merged})
    plain.unlink()


def _rewrite(source: Path, target: Path, change: dict[str, bytes | None]) -> None:
    """Copy the zip with parts replaced (or dropped, for None)."""
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w") as copy:
        for item in original.infolist():
            if item.filename not in change:
                copy.writestr(item, original.read(item))
            elif change[item.filename] is not None:
                copy.writestr(item.filename, change[item.filename])  # type: ignore[arg-type]


DAMAGE = [
    "no_types", "no_workbook", "no_sheet", "no_rels", "bad_sheet_xml", "bad_workbook_xml",
    "garbage_sheet", "empty_sheet_part", "wrong_types", "duplicate_sheets", "no_dimension",
    "bad_styles", "huge_row", "truncated", "notzip", "empty",
]  # fmt: skip


def damaged_copies(source: Path, directory: Path) -> dict[str, Path]:
    """Damaged versions of a workbook, by what is wrong with them (see `DAMAGE`)."""
    data = source.read_bytes()
    with zipfile.ZipFile(source) as package:
        names = package.namelist()
        sheet = next(name for name in names if name.startswith("xl/worksheets/sheet"))
        sheet_xml = package.read(sheet)
        workbook_xml = package.read("xl/workbook.xml")
    made: dict[str, Path] = {}

    def add(label: str, change: dict[str, bytes | None]) -> None:
        made[label] = directory / f"{source.stem}_{label}.xlsx"
        _rewrite(source, made[label], change)

    add("no_types", {"[Content_Types].xml": None})
    add("no_workbook", {"xl/workbook.xml": None})
    add("no_sheet", {sheet: None})
    add("no_rels", {"xl/_rels/workbook.xml.rels": None})
    add("bad_sheet_xml", {sheet: sheet_xml[: len(sheet_xml) // 2]})
    add("bad_workbook_xml", {"xl/workbook.xml": workbook_xml[:-30] + b"<<<"})
    add("garbage_sheet", {sheet: b"\x00\xff not xml at all"})
    add("empty_sheet_part", {sheet: b""})
    add("wrong_types", {"[Content_Types].xml": b'<?xml version="1.0"?><Types xmlns="x"/>'})
    add(
        "duplicate_sheets",
        {"xl/workbook.xml": re.sub(rb'(<sheet [^>]*?name=")[^"]*', rb"\1Dup", workbook_xml)},
    )
    add("no_dimension", {sheet: _drop(sheet_xml, b"<dimension", b"/>")})
    add("bad_styles", {"xl/styles.xml": b"<styleSheet/>"} if "xl/styles.xml" in names else {})
    add("huge_row", {sheet: sheet_xml.replace(b"<row ", b'<row r="2000000" ', 1)})
    for label, content in [
        ("truncated", data[: len(data) // 2]),
        ("notzip", b"PK\x03\x04 this is not a workbook"),
        ("empty", b""),
    ]:
        made[label] = directory / f"{source.stem}_{label}.xlsx"
        made[label].write_bytes(content)
    return made


def _drop(xml: bytes, start: bytes, end: bytes) -> bytes:
    first = xml.find(start)
    if first < 0:
        return xml
    return xml[:first] + xml[xml.index(end, first) + len(end) :]


# ---- schema-driven arguments -------------------------------------------------------------

REFS = [
    "A1", "B2", "A1:B2", "B2:A1", "A:A", "1:1", "A:C", "2:5", "A1:XFD1048576", "A:XFD", "Z99",
    "XFD1048576", "XFE1", "A1048577", "A0", "", "A1:A", "1:1048577", "Data!A1", "'Data'!A1:B2",
    "$A$1:$B$2", "a1", "A1,B2", "Sales", "Sales[#All]", "Rate", "ünï", "#REF!", "A1:B2:C3",
    " A1 ", "G1", "J1", "A1:E41", "A1:E10", "C2:C41", "J1:K2", "Hidden!A1", "Nope!A1", "A1:A1",
    "A2:A2000000", "1:1000000", "C:C", "Data!A:E", "Data!A1:E41", "'Very Hidden'!A1",
]  # fmt: skip
SHEETS = [
    "Data", "Hidden", "Very Hidden", "Locked", "Empty", "Sheet1", "Nope", "", "x" * 40,
    "<b>x</b>", "Ünï😀", "=1+1", "Data ", "data", "[x]", "a/b", "History", "Sheet",
]  # fmt: skip
GOOD_CELLS = ["A1", "B2", "D1", "G1", "E5", "J1", "F1", "A2", "C3"]
GOOD_REFS = [
    "A1", "B2", "A1:B2", "C2:C41", "A1:E41", "A1:E10", "A2:E41", "D1", "G1", "J1:K2", "A:A",
    "1:1", "B:C", "Data!A1:E41", "A1:C3", "E5", "F1:F10",
]  # fmt: skip
GOOD_NAMES = ["Sales", "Rate", "Region", "Product", "Units", "Price", "Chart 1", "Picture 1",
              "PivotTable1", "Slicer_Region", "Module1", "Total", "x"]  # fmt: skip
GOOD_FORMULAS = ["=A1+1", "=SUM(C2:C41)", "Data!$A$1", "Data!A1:B2", "=Rate*2", "Data!$C$2:$C$41",
                 "=IF(A1>1,1,0)", "1", "=A2"]  # fmt: skip
STRINGS = [
    "", " ", "a", "x" * 300, "x" * 40_000, "Ünï😀", "<b>&amp;</b>", "=1+1", "=SUM(", "\n", "A1",
    "name", "[x]", "{}", "%s", "0", "-1", "1e400", "=" + "(" * 70 + "1" + ")" * 70, "=A1+A1",
    "=HYPERLINK(\"http://x\",\"y\")", "'sheet'!A1", "Data!A1:B2", "\\\\host\\share\\x.xlsx",
    "http://example.com", "#Sheet!A1", "True", "2024-01-01", "1/0", "#N/A", "x\u202ey", "\t",
]  # fmt: skip
INTS = [0, 1, -1, 2, 3, 5, 10, 100, 1000, 100_000, 1_000_000, 2**31, 2**53, -(2**31), 255, 256]
NUMBERS = [0.0, 1.0, -1.0, 0.5, 1e308, -1e308, 1e-308, 255.5, 1e15, 3.14159]
NAMES = ["Sales", "Rate", "Chart 1", "Picture 1", "Table1", "PivotTable1", "Slicer_Region", "x",
         "", "A1", "R1C1", "1abc", "a b", "<x>", "x" * 300, "Ünï", "Data"]  # fmt: skip
WEIRD_PATHS = [
    "", "missing.xlsx", "../x.xlsx", "x.txt", "sub", "nul.xlsx", "C:\\Windows\\win.ini",
    "\\\\host\\share\\a.xlsx", "a.xlsx:stream", "x\x00y.xlsx", "CON.xlsx", ".",
]  # fmt: skip
FORMULAS = ["=1+1", "=INDIRECT(\"A1\")", "Data!$A$1", "Data!A1:B2", "=SUM(Data!C2:C41)", "#REF!",
            "1", "\"text\"", "=A1", "OFFSET(Data!A1,1,1)", "Data!$A:$A", "=NA()", "Rate", "",
            "=" + "ABS(" * 70 + "1" + ")" * 70, "Data!1:1"]  # fmt: skip


class Arguments:
    """Random arguments for a tool, following its JSON schema, drawn from pools of nasty values."""

    def __init__(self, seed: int, workbook: str, files: dict[str, str], sheets: list[str]) -> None:
        self.random = random.Random(seed)
        self.workbook = workbook
        self.files = files
        self.sheets = sheets

    def make(self, schema: dict[str, Any]) -> dict[str, Any]:
        arguments: dict[str, Any] = {}
        required = set(schema.get("required", []))
        for name, sub in schema.get("properties", {}).items():
            if name in required or self.random.random() < 0.5:
                arguments[name] = self.value(sub, schema, name)
        return arguments

    def value(self, sub: dict[str, Any], root: dict[str, Any], name: str = "") -> Any:
        pick = self.random
        if "$ref" in sub:
            sub = root["$defs"][sub["$ref"].split("/")[-1]]
        for key in ("anyOf", "oneOf"):
            if key in sub:
                return self.value(pick.choice(sub[key]), root, name)
        if "enum" in sub:
            return pick.choice(sub["enum"])
        if "const" in sub:
            return sub["const"]
        kind = sub.get("type")
        if isinstance(kind, list):
            kind = pick.choice(kind)
        match kind:
            case "null":
                return None
            case "boolean":
                return pick.random() < 0.5
            case "integer":
                return self._number(sub, pick.choice(INTS))
            case "number":
                return self._number(sub, pick.choice(NUMBERS + [float(i) for i in INTS]))
            case "string":
                if sub.get("format") in ("date", "date-time"):
                    return pick.choice(["2024-01-01", "2024-02-30", "0001-01-01", "9999-12-31"])
                return self.text(name)
            case "array":
                items = sub.get("items", {})
                count = pick.choice([0, 1, 1, 2, 3, 20])
                return [self.value(items, root, name) for _ in range(count)]
            case "object":
                return self._object(sub, root)
        return pick.choice(STRINGS)

    def _number(self, sub: dict[str, Any], value: float) -> float:
        """Mostly a number the schema allows, often at its edge, to get past validation."""
        low, high = sub.get("minimum", sub.get("exclusiveMinimum")), sub.get("maximum")
        step = 1 if sub.get("type") == "integer" else 0.5
        if "exclusiveMinimum" in sub:
            low = sub["exclusiveMinimum"] + step
        if self.random.random() < 0.85:
            if low is not None and high is not None and self.random.random() < 0.4:
                return self.random.choice([low, high, (low + high) // 2])
            if low is not None:
                value = max(value, low)
            if high is not None:
                value = min(value, high)
        return value

    def _object(self, sub: dict[str, Any], root: dict[str, Any]) -> dict[str, Any]:
        if "properties" in sub:
            return self.make({**sub, "$defs": root.get("$defs", {})})
        extra = sub.get("additionalProperties")
        if isinstance(extra, dict):
            return {
                self.random.choice(["A", "B", "Rate", "x", ""]): self.value(extra, root)
                for _ in range(self.random.choice([0, 1, 2]))
            }
        return {}

    def text(self, name: str) -> str:
        """A string for the parameter: mostly a plausible one, so that calls get past validation."""
        pick = self.random
        plausible = pick.random() < 0.8
        if name == "path":
            return self.workbook if pick.random() < 0.93 else pick.choice(WEIRD_PATHS)
        if name in ("csv_path", "image_path", "directory"):
            good = [v for k, v in self.files.items() if k != "base64"] + ["out.csv", "."]
            return pick.choice(good if plausible else WEIRD_PATHS)
        if name == "content_base64":
            return pick.choice([self.files["base64"], "", "!!!", "UEsDBA==", "A" * 10_000])
        if name in ("sheet", "to_sheet"):
            return pick.choice(self.sheets if plausible else SHEETS)
        if name == "new_name":
            return pick.choice(["New", "Renamed", "Copy of Data"] if plausible else SHEETS)
        if name in ("name", "module", "field", "caption"):
            return pick.choice(GOOD_NAMES if plausible else NAMES + STRINGS[:12])
        if name == "query":
            return pick.choice(["North", "P1", "Region", "1", "a"] if plausible else STRINGS)
        if name in ("refers_to", "formula", "formula1", "formula2", "code"):
            return pick.choice(GOOD_FORMULAS if plausible else FORMULAS + STRINGS)
        if name in ("at", "cell", "changing_cell", "start"):
            return pick.choice(GOOD_CELLS if plausible else REFS)
        if name in ("range", "source", "categories"):
            return pick.choice(GOOD_REFS if plausible else REFS)
        return pick.choice(GOOD_NAMES + GOOD_REFS if plausible else REFS + STRINGS + NAMES)


def sample_files(directory: Path) -> dict[str, str]:
    """Input files the tools may be pointed at, by kind."""
    (directory / "in.csv").write_text('a,b\n1,"x,y"\n2025-01-02,3.5\n', encoding="utf-8")
    (directory / "empty.csv").write_text("")
    (directory / "quotes.csv").write_text('"' + "," * 200_000 + '"\n')
    (directory / "in.png").write_bytes(PNG)
    (directory / "bad.png").write_bytes(b"not an image")
    return {
        "csv": "in.csv",
        "empty": "empty.csv",
        "quotes": "quotes.csv",
        "png": "in.png",
        "badpng": "bad.png",
        "base64": base64.b64encode(PNG).decode(),
    }


def sheet_names(path: Path) -> list[str]:
    """The sheets of a workbook, or those of the rich workbook when it cannot be read."""
    with path.open("rb") as handle:  # openpyxl leaves a file it opened itself open on error
        try:
            workbook = load_workbook(handle, read_only=True)
        except Exception:
            return ["Data", "Hidden", "Very Hidden", "Locked", "Empty"]
        names = workbook.sheetnames or ["Sheet1"]
        workbook.close()
        return names
