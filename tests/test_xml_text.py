"""Client text that reaches XML: names with markup characters keep the package well-formed."""

import re
from collections.abc import Callable, Coroutine
from pathlib import Path
from typing import Any

import pytest
from PIL import Image

from tests.conftest import ToolCall
from tests.package_support import assert_namespaces_declared, read_parts
from tests.slicer_support import SALES, TABLE, add, build_book, parts_of

pytestmark = pytest.mark.anyio

MARKUP = ["a\"<>&'b", '" /><x:evil xmlns:x="u"/><a b="', "]]>&amp;&#10;&lt;", "tab\there\nnl"]
CONTROL = ["a\x01b", "a\x0bb", "a\ud800b", "a￾b", "a\x00b"]
CallOutcome = Callable[..., Coroutine[Any, Any, tuple[bool, str]]]


@pytest.fixture
def attempt(client: Any) -> CallOutcome:
    async def run(tool: str, /, **arguments: Any) -> tuple[bool, str]:
        result = await client.call_tool(tool, arguments)
        text = result.content[0].text if result.content else ""
        assert "Unexpected error" not in text, text
        return result.is_error, text

    return run


def assert_well_formed(path: str | Path) -> None:
    assert_namespaces_declared(read_parts(Path(path)))


@pytest.mark.parametrize("text", CONTROL)
async def test_control_characters_are_refused_in_every_argument(
    call: ToolCall, call_error: ToolCall, sample: Path, text: str
) -> None:
    cases = [
        ("create_workbook", {"path": "n.xlsx", "sheets": [text]}),
        ("create_sheet", {"path": "sales.xlsx", "new_name": text}),
        ("rename_sheet", {"path": "sales.xlsx", "sheet": "Report", "new_name": text}),
        ("set_defined_name", {"path": "sales.xlsx", "name": text, "refers_to": "Data!A1"}),
        ("set_note", {"path": "sales.xlsx", "sheet": "Data", "cell": "A1", "text": text}),
        (
            "set_note",
            {"path": "sales.xlsx", "sheet": "Data", "cell": "A1", "text": "x", "author": text},
        ),
        ("set_table", {"path": "sales.xlsx", "sheet": "Data", "range": "A1:D5", "name": text}),
        (
            "insert_image",
            {
                "path": "sales.xlsx",
                "sheet": "Data",
                "image_path": "i.png",
                "at": "F2",
                "name": text,
            },
        ),
        (
            "add_slicer",
            {
                "path": "sales.xlsx",
                "sheet": "Data",
                "target": TABLE,
                "field": "Region",
                "at": "F2",
                "name": text,
            },
        ),
        (
            "add_slicer",
            {
                "path": "sales.xlsx",
                "sheet": "Data",
                "target": TABLE,
                "field": "Region",
                "at": "F2",
                "caption": text,
            },
        ),
        (
            "create_chart",
            {
                "path": "sales.xlsx",
                "sheet": "Data",
                "chart_type": "column",
                "at": "F2",
                "source": "A1:C5",
                "name": text,
            },
        ),
        (
            "create_chart",
            {
                "path": "sales.xlsx",
                "sheet": "Data",
                "chart_type": "column",
                "at": "F2",
                "source": "A1:C5",
                "options": {"title": text},
            },
        ),
        (
            "set_workbook_settings",
            {"path": "sales.xlsx", "settings": {"doc_properties": {"company": text}}},
        ),
    ]
    for tool, arguments in cases:
        assert "control characters" in await call_error(tool, **arguments), (tool, arguments)


@pytest.mark.parametrize("name", MARKUP)
async def test_slicer_names_and_captions_are_escaped(
    call: ToolCall, files: Path, attempt: CallOutcome, name: str
) -> None:
    path = await build_book(call, files)
    await add(call, path, sheet="Data", target=TABLE, field="Region", at="G2", name=name)
    await add(call, path, sheet="Pivot", at="E3", field="Product", name=name + "2", caption=name)
    await add(call, path, sheet="Data", target=TABLE, field="Product", at="G20", caption=name)
    assert_well_formed(path)
    listed = await call("describe_sheet", path=path, sheet="Data")
    names = {item["name"] for item in listed["slicers"]}
    assert name in names
    await call("delete_slicer", path=path, sheet="Data", name=name)
    await call("copy_sheet", path=path, sheet="Pivot", new_name="Copy")
    await call("copy_sheet", path=path, sheet="Data", new_name="Copy2")
    assert_well_formed(path)
    after = {
        item["name"] for item in (await call("describe_sheet", path=path, sheet="Copy"))["slicers"]
    }
    assert len(after) == 1


@pytest.mark.parametrize("name", MARKUP)
async def test_slicer_over_markup_field_and_sheet_names(
    call: ToolCall, files: Path, name: str
) -> None:
    path = str(files / "markup.xlsx")
    sheet = sheet_name(name)
    await call("create_workbook", path=path, sheets=[sheet, "Pivot"])
    header = [f"{h}{name}" for h in SALES[0]]
    await call("write_range", path=path, sheet=sheet, at="A1", rows=[header, *SALES[1:]])
    await call("set_table", path=path, sheet=sheet, range="A1:D9", name="Sales")
    await call(
        "create_pivot_table",
        path=path,
        source=f"'{sheet.replace(chr(39), chr(39) * 2)}'!A1:D9",
        row_fields=[header[1]],
        value_fields=[{"field": header[3]}],
        sheet="Pivot",
        at="A3",
        name=name,
    )
    await call(
        "add_slicer",
        path=path,
        sheet=sheet,
        target={"sheet": sheet, "name": "Sales"},
        field=header[0],
        at="G2",
    )
    await call(
        "add_slicer",
        path=path,
        sheet="Pivot",
        target={"sheet": "Pivot", "name": name},
        field=header[1],
        at="G2",
    )
    await call("rename_sheet", path=path, sheet=sheet, new_name="Renamed")
    await call("copy_sheet", path=path, sheet="Pivot", new_name="Copy")
    assert_well_formed(path)


@pytest.mark.parametrize("name", MARKUP)
async def test_object_names_and_titles_are_escaped(
    call: ToolCall, files: Path, sample: Path, attempt: CallOutcome, name: str
) -> None:
    Image.new("RGB", (10, 10), "red").save(files / "i.png")
    book = {"path": "sales.xlsx"}
    await call("create_sheet", **book, new_name=sheet_name(name))
    await call("set_table", **book, sheet="Data", range="A1:D5", name="Tbl")
    await call("insert_image", **book, sheet="Report", image_path="i.png", at="H2", name=name)
    await call("set_note", **book, sheet="Report", cell="A1", text=name, author=name)
    await call("set_defined_name", **book, name="Region_1", refers_to="Data!$A$1:$A$5")
    for chart_type in ("column", "pie", "waterfall", "histogram", "box_whisker"):
        await call(
            "create_chart",
            **book,
            sheet="Report",
            chart_type=chart_type,
            at="J2" if chart_type == "column" else f"{'LNPRT'[len(chart_type) % 5]}20",
            source="Data!B1:C5" if chart_type in ("column", "pie") else "Data!C1:C5",
            name=name + chart_type,
            options={"title": name, "x_axis": {"title": name}}
            if chart_type == "column"
            else {"title": name},
        )
    await call(
        "create_chart",
        **book,
        sheet="Report",
        chart_type="line",
        at="A30",
        series=[{"values": "Data!C2:C5", "name": name}],
        options={"title": name, "y_axis": {"title": name, "number_format": name}},
    )
    await call(
        "write_range",
        **book,
        sheet="Report",
        at="A10",
        rows=[["x"]],
        links=[{"cell": "A10", "target": "https://example.com/?a=1&b=<2>", "tooltip": name}],
    )
    await call(
        "set_workbook_settings",
        **book,
        settings={
            "doc_properties": {"title": name, "author": name, "company": name, "keywords": name}
        },
    )
    await call(
        "format_range",
        **book,
        sheet="Data",
        range="A1:B2",
        style={"number_format": name + "0", "font_name": name},
    )
    outcome = [
        await attempt(
            "add_conditional_format",
            **book,
            sheet="Data",
            range="C2:C5",
            rule={"type": "contains_text", "text": name, "fill_color": "#FF0000"},
        ),
        await attempt(
            "add_data_validation",
            **book,
            sheet="Data",
            range="C2:C5",
            rule={
                "type": "list",
                "options": [name, "b"],
                "prompt_title": name[:32],
                "prompt": name,
                "error_title": name[:32],
                "error_message": name,
            },
        ),
        await attempt(
            "set_sheet_layout",
            **book,
            sheet="Data",
            layout={
                "print_setup": {"header": {"left": name, "right": name}, "footer": {"center": name}}
            },
        ),
        await attempt(
            "set_table",
            **book,
            sheet="Data",
            name="Tbl",
            options={"columns": [{"name": "Region", "total_label": name}]},
        ),
    ]
    assert_well_formed(files / "sales.xlsx")
    assert all(isinstance(failed, bool) for failed, _ in outcome)
    listed = await call("describe_sheet", **book, sheet="Report")
    assert any(note["text"] == name for note in listed["notes"])


@pytest.mark.parametrize("name", MARKUP)
async def test_pivot_names_and_fields_are_escaped(call: ToolCall, files: Path, name: str) -> None:
    path = str(files / "pivot.xlsx")
    await call("create_workbook", path=path, sheets=["Data", "Pivot"])
    header = [f"{h}{name}" for h in SALES[0]]
    await call("write_range", path=path, sheet="Data", at="A1", rows=[header, *SALES[1:]])
    await call(
        "create_pivot_table",
        path=path,
        source="Data!A1:D9",
        row_fields=[header[1]],
        column_fields=[header[0]],
        value_fields=[{"field": header[3], "number_format": f'0"{name[:3]}"'.replace('""', "")}],
        calculated_fields=[
            {"name": "Calc" + name, "formula": "='" + header[3].replace("'", "''") + "'*2"}
        ],
        sheet="Pivot",
        at="A3",
        name=name,
    )
    await call(
        "add_slicer",
        path=path,
        sheet="Pivot",
        target={"sheet": "Pivot", "name": name},
        field=header[0],
        at="L3",
        name=name,
    )
    await call(
        "add_slicer",
        path=path,
        sheet="Pivot",
        target={"sheet": "Pivot", "name": name},
        field=header[2],
        at="L20",
        timeline={"level": "months"},
    )
    assert_well_formed(path)


def sheet_name(text: str) -> str:
    return re.sub(r"[\[\]:*?/\\n\t]", "", text).strip("'")[:31] or "S"


@pytest.mark.parametrize("name", ["a\"<>&'b", 'x &amp; <y> "z"', "R&D"])
async def test_sheet_names_with_markup_in_every_kind_of_reference(
    call: ToolCall, files: Path, attempt: CallOutcome, name: str
) -> None:
    path = str(files / "refs.xlsx")
    await call("create_workbook", path=path, sheets=["Data", "Other"])
    await call("write_range", path=path, sheet="Data", at="A1", rows=SALES)
    await call("set_table", path=path, sheet="Data", range="A1:D9", name="Sales")
    await call("add_slicer", path=path, sheet="Data", target=TABLE, field="Region", at="G2")
    await call("add_sparklines", path=path, sheet="Data", range="F2:F4", source="Data!D2:D4")
    await call("add_sparklines", path=path, sheet="Data", range="H2:H4", source="Data!B2:D4")
    for rule in (
        {"type": "data_bar", "colors": ["#638EC6"], "min_value": "=Data!$D$2", "max_value": "99"},
        {"type": "icon_set", "icons": ["3_arrows"]},
    ):
        await attempt("add_conditional_format", path=path, sheet="Data", range="D2:D9", rule=rule)
    for chart_type in ("column", "waterfall"):
        await call(
            "create_chart",
            path=path,
            sheet="Data",
            chart_type=chart_type,
            at="J2" if chart_type == "column" else "J20",
            source="Data!C1:D9" if chart_type == "column" else "Data!D1:D9",
        )
    await call("set_defined_name", path=path, name="Totals", refers_to="Data!$D$2:$D$9")
    await call("rename_sheet", path=path, sheet="Data", new_name=name)
    assert_well_formed(path)
    await call("copy_sheet", path=path, sheet=name, new_name=name + "2")
    await call("rename_sheet", path=path, sheet=name + "2", new_name="Plain")
    await call("rename_sheet", path=path, sheet=name, new_name="Plain2")
    assert_well_formed(path)


async def test_slicer_names_follow_excels_limits(
    call: ToolCall, call_error: ToolCall, files: Path
) -> None:
    path = await build_book(call, files)
    for bad in ("", "   ", "x" * 256):
        message = await call_error(
            "add_slicer", path=path, sheet="Data", target=TABLE, field="Region", at="G2", name=bad
        )
        assert "1 to 255 characters" in message
    message = await call_error(
        "add_slicer",
        path=path,
        sheet="Data",
        target=TABLE,
        field="Region",
        at="G2",
        caption="y" * 256,
    )
    assert "caption" in message
    await add(call, path, sheet="Data", target=TABLE, field="Region", at="G2", name="x" * 255)
    assert "exists" in await call_error(
        "add_slicer",
        path=path,
        sheet="Data",
        target=TABLE,
        field="Product",
        at="G20",
        name="X" * 255,
    )


async def test_cache_names_keep_what_excel_keeps(call: ToolCall, files: Path) -> None:
    path = str(files / "cache.xlsx")
    await call("create_workbook", path=path)
    fields = ["a.b-c d", "9lives", "é-ü ß", r"a\b?c", "Q&A <x>"]
    await call("write_range", path=path, sheet="Sheet1", at="A1", rows=[fields, *[["v"] * 5] * 3])
    await call("set_table", path=path, sheet="Sheet1", range="A1:E4", name="T")
    for index, field in enumerate(fields):
        await call(
            "add_slicer",
            path=path,
            sheet="Sheet1",
            target={"sheet": "Sheet1", "name": "T"},
            field=field,
            at=f"{'GHIJK'[index]}2",
        )
    names = re.findall(
        r'name="([^"]*)" sourceName',
        "".join(data.decode() for part, data in parts_of(path).items() if "slicerCaches/" in part),
    )
    assert names == [
        "Slicer_a.b_c_d",
        "Slicer_9lives",
        "Slicer_é_ü_ß",
        r"Slicer_a\b?c",
        "Slicer_Q_A__x",
    ]
