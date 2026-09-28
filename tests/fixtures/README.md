# Test fixtures

Workbooks that tests cannot build in code.

| File | Source | License |
| --- | --- | --- |
| `vbaProject.bin` | [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) `examples/vbaProject.bin` | BSD-2-Clause, Copyright John McNamara |
| `macro01.xlsm` | [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) `xlsxwriter/test/comparison/xlsx_files/macro01.xlsm` | BSD-2-Clause, Copyright John McNamara |

Both were saved by Excel and contain small VBA projects. Tests only parse them; the
macros are never run.
