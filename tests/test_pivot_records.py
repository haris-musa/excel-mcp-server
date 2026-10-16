import datetime as dt
import io
import zipfile
from xml.etree import ElementTree

from openpyxl.pivot.fields import DateTimeField, Index, Missing, Number, Text
from openpyxl.pivot.record import Record, RecordList
from openpyxl.xml.constants import SHEET_MAIN_NS

from excel_mcp.operations.pivot_records import PivotRecords, SharedIndex
from excel_mcp.package.lazy_records import LazyRecords


def _written(records: RecordList) -> str:
    buffer = io.BytesIO()
    manifest: list[RecordList] = []
    with zipfile.ZipFile(buffer, "w") as archive:
        records._write(archive, manifest)  # pyright: ignore[reportAttributeAccessIssue]
        assert manifest == [records]
        return ElementTree.canonicalize(archive.read(records.path[1:]))


def test_records_are_written_as_openpyxl_writes_them() -> None:
    moment = dt.datetime(2024, 3, 5, 7, 9, 11, 250000)
    rows = [
        (SharedIndex(2), 'a<b>&"c"\n\t', 1.5, 7, moment, None),
        (SharedIndex(0), "", -0.1, 2**70, dt.datetime(2020, 1, 1), None),
    ]
    expected = RecordList(
        r=[
            Record(
                _fields=[
                    Index(v=int(row[0])),
                    Text(v=row[1]),
                    Number(v=row[2]),
                    Number(v=row[3]),
                    DateTimeField(v=row[4]),
                    Missing(),
                ]
            )
            for row in rows
        ]
    )
    assert PivotRecords(rows).count == expected.count == 2
    assert _written(PivotRecords(rows)) == _written(expected)


RECORDS = (
    f'<pivotCacheRecords xmlns="{SHEET_MAIN_NS}" count="3">'
    '<r><s v="a&amp;b" /><x v="0" /><n v="1.5" /></r>'
    '<r><m /><x v="1" /><d v="2024-01-02T00:00:00" /></r>'
    '<r><s v="A&#10;B" /><x v="1" /><b v="1" /></r>'
    "</pivotCacheRecords>"
)


def test_lazy_records_hand_out_one_column() -> None:
    records = LazyRecords(RECORDS.encode())
    assert records.count == 3
    first, second, third = records.column(0)
    assert (first.v, third.v) == ("a&b", "A\nB")
    assert isinstance(second, Missing)
    assert [item.v for item in records.column(2)] == [1.5, dt.datetime(2024, 1, 2), True]


def test_lazy_records_replace_one_column_and_keep_the_rest() -> None:
    records = LazyRecords(RECORDS.encode())
    records.share(0, [2, 0, 1])
    expected = RECORDS.replace('<s v="a&amp;b" />', '<x v="2" />')
    expected = expected.replace("<m />", '<x v="0" />').replace('<s v="A&#10;B" />', '<x v="1" />')
    assert _written(records) == ElementTree.canonicalize(expected)
