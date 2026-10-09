# Test fixtures

Workbooks that tests cannot build in code.

| File | Source | License |
| --- | --- | --- |
| `vbaProject.bin` | [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) `examples/vbaProject.bin` | BSD-2-Clause, Copyright John McNamara |
| `macro01.xlsm` | [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) `xlsxwriter/test/comparison/xlsx_files/macro01.xlsm` | BSD-2-Clause, Copyright John McNamara |
| `excel_sparklines.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_chartex.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_slicers.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_shapes.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_shapes.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_objects.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_controls.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `excel_dynamic_arrays.xlsx` | Created in Microsoft Excel (COM) for this project | Same as the project |
| `formula_golden.json` | Results recorded from Microsoft Excel by `scripts/excel_golden.py` | Same as the project |

`macro01.xlsm` and `vbaProject.bin` were saved by Excel and contain small VBA projects. Tests
only parse them; the macros are never run.

The `excel_*.xlsx` files hold what openpyxl cannot model, as Excel itself saved it, so that
tests check the package layer against the real thing: sparklines with extended conditional
formats and data validation; a waterfall, histogram, treemap and funnel chart with a text box,
shape and picture; shapes, a text box, a connector, a group and a picture fill; an embedded object, a protected range and an ignored error; table, PivotTable and timeline slicers; form controls with a note; and the
dynamic array formulas `FILTER`, `SORT`, `SEQUENCE`, `UNIQUE` and `SUM(range*2)`. The local
path and user name that Excel records were removed.
