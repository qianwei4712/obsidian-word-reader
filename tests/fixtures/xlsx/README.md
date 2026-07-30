# XLSX validation fixture catalog

The XLSX fixtures are generated in `tests/xlsxFixture.ts` so their OOXML
structure remains reviewable and deterministic.

| Fixture | Coverage |
| --- | --- |
| `createRichXlsx()` | Valid workbook, workbook- and sheet-scoped named ranges, ignored dynamic/external names, shared and inline strings, merged cells, frozen rows/columns, hidden sheet, date and numeric formats, cached formulas, external-data formula marker, sparse `100,000`-row extent, external and internal hyperlinks, legacy rich-text comment with author, blank-cell comment discovery, package-local anchored PNG, cached bar chart, literal numeric `cellIs`, color-scale and data-bar rules, ignored external workbook relationship and ignored data connection |
| `createLargeDenseXlsx()` | Deterministic `20,000`-row/`80,000`-cell dense worksheet for streamed XML chunking, cancellation, maximum retained worksheet-data buffer, parse-time, peak-heap and virtual-grid budgets |
| `addOoxmlEntry()` variants | Macro, OLE, ActiveX and script-capable media rejection |
| Raw safety fixtures | Damaged ZIP, encrypted OLE container, ZIP64 marker, excessive file/entry/total limits and malicious compression ratio |

The fixture library remains generated, deterministic, package-local, and safe
to inspect. Public `.xlsx` view integration is covered separately; `.xlsm`
stays unregistered and macro-enabled variants are rejection-only fixtures.
