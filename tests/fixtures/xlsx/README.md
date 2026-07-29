# XLSX validation fixture catalog

The XLSX fixtures are generated in `tests/xlsxFixture.ts` so their OOXML
structure remains reviewable and deterministic.

| Fixture | Coverage |
| --- | --- |
| `createRichXlsx()` | Valid workbook, workbook- and sheet-scoped named ranges, ignored dynamic/external names, shared and inline strings, merged cells, frozen rows/columns, hidden sheet, date and numeric formats, cached formulas, external-data formula marker, sparse `100,000`-row extent, external and internal hyperlinks, raster image, ignored external workbook relationship and ignored data connection |
| `addOoxmlEntry()` variants | Macro, OLE, ActiveX and script-capable media rejection |
| Raw safety fixtures | Damaged ZIP, encrypted OLE container, ZIP64 marker, excessive file/entry/total limits and malicious compression ratio |

No fixture is opened through an Obsidian view in the `2.5.x` line. The sample
library exercises only the unregistered parser and virtual-grid research
modules.
