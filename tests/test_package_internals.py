"""The building blocks of the package layer, and the API that later features build on."""

import re
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_mcp.config import Limits
from excel_mcp.errors import WorkbookError
from excel_mcp.package import Link, Part, state_of, vml
from excel_mcp.package import extensions as ext
from excel_mcp.package.metadata import Metadata
from excel_mcp.package.opc import ContentTypes, Rel, parse_rels, resolve, write_rels
from excel_mcp.package.patch import insertions, splice
from excel_mcp.package.scan import (
    relationship_ids,
    remap_relationship_ids,
    scan,
    with_namespaces,
)
from excel_mcp.paths import PathPolicy
from excel_mcp.workspace import Workspace
from tests.package_support import assert_package_is_consistent, read_parts, relationships, text

R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def test_scan_lists_children_with_their_exact_text() -> None:
    data = (
        b'<?xml version="1.0"?><root xmlns="u" xmlns:p="v"><a x="1"/><p:b>t<c/></p:b><d></d></root>'
    )
    document = scan(data)

    assert [(c.local, c.qname, document.raw(c)) for c in document.children] == [
        ("a", "a", '<a x="1"/>'),
        ("b", "p:b", "<p:b>t<c/></p:b>"),
        ("d", "d", "<d></d>"),
    ]
    assert document.namespaces == {"": "u", "p": "v"}
    assert data[document.close_tag :] == b"</root>"


def test_scan_can_jump_over_the_sheet_data() -> None:
    data = b'<w xmlns="u"><top/><sheetData><row><c/></row></sheetData><a/><b x="1"/></w>'
    document = scan(data, "sheetData")

    assert [document.raw(c) for c in document.children] == ["<a/>", '<b x="1"/>']
    assert data[document.children[0].start :].startswith(b"<a/>")


def test_scan_rejects_documents_with_a_doctype() -> None:
    with pytest.raises(WorkbookError, match="DOCTYPE"):
        scan(b'<!DOCTYPE x [<!ENTITY a "b">]><x/>')
    with pytest.raises(WorkbookError, match="not well-formed"):
        scan(b"<x><y></x>")


def test_a_fragment_carries_the_namespaces_it_needs() -> None:
    scope = {"x14": "urn:x14", "xr": "urn:xr", "r": R, "mc": "urn:mc"}

    assert with_namespaces('<ext><x14:a xr:uid="1"/></ext>', scope) == (
        '<ext xmlns:x14="urn:x14" xmlns:xr="urn:xr"><x14:a xr:uid="1"/></ext>'
    )
    already = '<x14:a xmlns:x14="urn:other"/>'
    assert with_namespaces(already, scope) == already
    assert 'xmlns:x14="urn:x14"' in with_namespaces('<mc:Choice Requires="x14"/>', scope)


def test_relationship_ids_are_found_and_renumbered_by_attribute_not_by_text() -> None:
    xml = f'<c r:id="rId1" id="rId1" x="rId2"><d xmlns:q="{R}" q:embed="rId2"/>'
    xml += '<v o:relid="rId3"/></c>'
    scope = {"r": R}

    assert relationship_ids(xml, scope) == {"rId1", "rId2", "rId3"}
    remapped = remap_relationship_ids(xml, {"rId1": "rId9", "rId2": "rId8"}, scope)
    expected = f'<c r:id="rId9" id="rId1" x="rId2"><d xmlns:q="{R}" q:embed="rId8"/>'
    assert remapped == expected + "<v o:relid=" + '"rId3"/></c>'


def test_relationships_resolve_inside_the_package_only() -> None:
    assert (
        resolve("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml")
        == "xl/drawings/drawing1.xml"
    )
    assert resolve("xl/workbook.xml", "/docProps/app.xml") == "docProps/app.xml"
    assert resolve("", "xl/workbook.xml") == "xl/workbook.xml"
    with pytest.raises(WorkbookError, match="leaves the package"):
        resolve("xl/workbook.xml", "../../outside.xml")


def test_relationships_survive_being_read_and_written() -> None:
    rels = [
        Rel("rId1", R + "image", "xl/media/a&b.png"),
        Rel("rId2", R + "hyperlink", "https://example.com/?a=1&b=2", external=True),
    ]

    assert parse_rels("xl/drawings/drawing1.xml", write_rels(rels)) == rels


def test_content_types_prefer_a_default_for_media() -> None:
    types = ContentTypes(
        b'<Types xmlns="x"><Default Extension="xml" ContentType="application/xml"/></Types>'
    )
    types.declare("xl/media/a.png", "image/png", by_default=True)
    types.declare("xl/media/b.png", "image/png", by_default=True)
    types.declare("xl/slicers/slicer1.xml", "application/vnd.ms-excel.slicer+xml")

    assert types.defaults == {"xml": "application/xml", "png": "image/png"}
    assert types.overrides == {"xl/slicers/slicer1.xml": "application/vnd.ms-excel.slicer+xml"}


def test_metadata_gets_the_dynamic_array_entry_once() -> None:
    fresh = Metadata(None)
    assert fresh.dynamic_array_index() == 1
    assert fresh.dynamic_array_index() == 1
    assert fresh.xml.count("<metadataType ") == 1 and fresh.xml.count("<bk>") == 2

    other = (
        '<metadata xmlns="m"><metadataTypes count="1"><metadataType name="XLRICHVALUE"/>'
        '</metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk/></futureMetadata>'
        '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>'
    )
    merged = Metadata(other)
    assert merged.dynamic_array_index() == 1
    assert re.findall(r'<metadataType name="(\w+)"', merged.xml) == ["XLRICHVALUE", "XLDAPR"]
    assert merged.xml.index("<cellMetadata") < merged.xml.index("<valueMetadata")
    assert '<rc t="2" v="0"/>' in merged.xml and "valueMetadata count" in merged.xml


def test_references_in_extensions_follow_what_the_caller_decides() -> None:
    sparklines = (
        f'<ext uri="{ext.SPARKLINES}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines>'
        "<x14:sparkline><xm:f>Data!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline>"
        "<x14:sparkline><xm:f>Data!A2:C2</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline>"
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
    )
    validations = (
        f'<ext uri="{ext.DATA_VALIDATIONS}"><x14:dataValidations count="2" xmlns:xm="m">'
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Lists!$A$1</xm:f></x14:formula1>'
        "<xm:sqref>B2</xm:sqref></x14:dataValidation>"
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Lists!$A$2</xm:f></x14:formula1>'
        "<xm:sqref>B9</xm:sqref></x14:dataValidation></x14:dataValidations></ext>"
    )
    found = {ext.SPARKLINES: sparklines, ext.DATA_VALIDATIONS: validations}

    def update(reference: str, kind: str) -> str | None:
        if kind == "range" and reference == "D2":
            return None  # the cell was deleted
        if kind == "range":
            return reference.replace("2", "3") if reference == "B2" else reference
        return reference.replace("Data", "Moved")

    ext.rewrite_references(found, update)  # type: ignore[arg-type]

    assert found[ext.SPARKLINES].count("<x14:sparkline>") == 1
    assert "<xm:f>Moved!A1:C1</xm:f>" in found[ext.SPARKLINES]
    assert "<xm:sqref>B3</xm:sqref>" in found[ext.DATA_VALIDATIONS]
    assert 'count="2"' in found[ext.DATA_VALIDATIONS]


def test_a_deleted_sheet_takes_the_sparklines_that_read_it() -> None:
    sparklines = (
        f'<ext uri="{ext.SPARKLINES}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines>'
        "<x14:sparkline><xm:f>'My Sheet'!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline>"
        "<x14:sparkline><xm:f>Data!A1:C1</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline>"
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
    )
    found = {ext.SPARKLINES: sparklines}

    ext.forget_sheets(found, {"My Sheet"})
    assert found[ext.SPARKLINES].count("<x14:sparkline>") == 1

    ext.forget_sheets(found, {"Data"})
    assert ext.SPARKLINES not in found


def test_notes_are_numbered_above_the_shapes_that_other_parts_name() -> None:
    document = scan(
        b'<xml xmlns:v="v" xmlns:o="o" xmlns:x="x"><o:shapelayout/>'
        b'<v:shapetype id="_x0000_t201"/><v:shapetype id="_x0000_t202"/>'
        b'<v:shape id="_x0000_s1025" type="#_x0000_t201">'
        b'<x:ClientData ObjectType="Checkbox"/></v:shape>'
        b'<v:shape id="_x0000_s1026" type="#_x0000_t202">'
        b'<x:ClientData ObjectType="Note"/></v:shape></xml>'
    )
    kept = vml.preserved_shapes(document)
    written = (
        b'<xml><shapetype id="_x0000_t202"/><shape id="_x0000_s1026" type="#_x0000_t202"/>'
        b'<shape id="_x0000_s1027" type="#_x0000_t202"/></xml>'
    )

    merged = vml.merge(written, kept, document.namespaces).decode()

    types = re.findall(r'<(?:v:)?shapetype\b[^>]*\bid="(\w+)"', merged)
    assert types == ["_x0000_t202", "_x0000_t201"]
    assert re.findall(r'<shape id="(_x0000_s\d+)"', merged) == ["_x0000_s1026", "_x0000_s1027"]
    assert re.findall(r'<v:shape\b[^>]*\bid="(_x0000_s\d+)"', merged) == ["_x0000_s1025"]
    assert 'ObjectType="Note"' not in merged


def test_new_content_is_placed_where_the_schema_wants_it() -> None:
    data = b'<w xmlns="u"><sheetData/><pageMargins/><drawing/><tableParts/></w>'
    document = scan(data, "sheetData")
    order = [
        "sheetData",
        "pageMargins",
        "drawing",
        "legacyDrawing",
        "picture",
        "tableParts",
        "extLst",
    ]

    edits = insertions(document, [("extLst", "<extLst/>"), ("picture", "<picture/>")], order)

    assert splice(data, edits) == (
        b'<w xmlns="u"><sheetData/><pageMargins/><drawing/><picture/><tableParts/><extLst/></w>'
    )


def test_features_can_add_parts_and_extensions_through_the_state(files: Path) -> None:
    """What a feature that writes sparklines, slicers or comments does, end to end."""
    book = files / "book.xlsx"
    Workbook().save(book)
    workspace = Workspace(PathPolicy([files]), Limits(), allow_macro_workbooks=False)
    part = Part("xl/custom/thing1.xml", "application/vnd.example+xml", b"<thing/>")
    shared = Part("xl/media/image1.png", "image/png", b"\x89PNG")
    part.links.append(Link("rId1", f"{R}/image", shared))

    with workspace.edit("book.xlsx") as workbook:
        sheet = state_of(workbook).sheet(workbook.worksheets[0])
        sheet.links.append(Link("rIdNew", "urn:example:rel", part))
        sheet.set_extension(
            "{00000000-0000-0000-0000-000000000000}",
            '<ext uri="{00000000-0000-0000-0000-000000000000}" xmlns:x="urn:x">'
            f'<x:ref r:id="rIdNew" xmlns:r="{R}"/></ext>',
        )
        sheet.set_element("phoneticPr", '<phoneticPr fontId="0"/>')
        state_of(workbook).workbook.set_extension(
            "{11111111-0000-0000-0000-000000000000}", "<ext/>"
        )
    parts = read_parts(book)
    sheet_xml = text(parts, "xl/worksheets/sheet1.xml")
    rel_id = re.search(r'<x:ref r:id="(\w+)"', sheet_xml)

    assert rel_id and relationships(parts, "xl/worksheets/sheet1.xml")[rel_id[1]][1] == part.name
    assert relationships(parts, part.name)["rId1"][1] == shared.name
    assert parts[part.name] == b"<thing/>" and parts[shared.name] == b"\x89PNG"
    assert sheet_xml.index("<phoneticPr") < sheet_xml.index("<pageMargins")
    assert sheet_xml.rindex("<extLst>") > sheet_xml.index("<pageMargins")
    assert "<extLst><ext/></extLst>" in text(parts, "xl/workbook.xml")
    assert_package_is_consistent(parts)
