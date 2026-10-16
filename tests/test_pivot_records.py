import datetime as dt
import io
import zipfile
from xml.etree import ElementTree

from openpyxl.pivot.fields import DateTimeField, Index, Missing, Number, Text
from openpyxl.pivot.record import Record, RecordList

from excel_mcp.operations.pivot_records import PivotRecords, SharedIndex


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
