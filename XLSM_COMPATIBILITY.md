# XLSM read-only compatibility decision

Version reviewed: `3.2.0`

Decision: Office Reader does not register `.xlsm` files in the `3.2.0`
release.

## Evidence

| Case | Expected result |
| --- | --- |
| Ordinary `.xlsx` package | Opens through the local read-only XLSX pipeline |
| Macro-enabled workbook content type | Rejected as active content |
| `vbaProject.bin`, macro sheet, dialog sheet, or ActiveX part | Rejected as active content |
| Embedded OLE object | Rejected |
| Script-capable media | Rejected |
| External workbook relationship or data connection | Never fetched or evaluated |
| Stored formulas and chart series | Display cached values only; never recalculate |

The automated suite covers the registered-extension boundary and the
macro-enabled content-type, VBA, ActiveX, OLE, script-media, external
relationship, damaged package, encryption, and package-limit paths.

## Rationale

An `.xlsm` extension communicates that active content may be present. Opening
such a file in a view that intentionally discards macros could create a false
impression that the complete workbook was inspected. The shared OOXML policy
also rejects macro-enabled content types before worksheet parsing, so merely
adding the extension would produce an inconsistent user experience.

Read-only `.xlsm` support can be reconsidered only after a dedicated corpus
covers signed and unsigned VBA projects, macro sheets, ActiveX controls,
external connections, malformed projects, and workbooks whose cached values
depend on macros. Macro code will not be executed.

# XLSM 只读兼容性结论

评估版本：`3.2.0`

结论：Office Reader 在 `3.2.0` 中不注册 `.xlsm`。

`.xlsm` 扩展名表示文件可能包含活动内容。如果阅读器静默丢弃宏后仍以完整工作簿
形式打开，容易让用户误以为已经检查了全部逻辑。当前统一 OOXML 防护也会在解析
工作表前拒绝宏启用内容类型、VBA、宏工作表、ActiveX、OLE 和脚本型媒体，因此
直接注册扩展名会产生不一致体验。

自动化样本已经覆盖宏启用内容类型、VBA/ActiveX/OLE/脚本媒体、外部关系、损坏、
加密和包上限。只有在补齐签名/未签名 VBA、宏工作表、外部连接、损坏工程以及
依赖宏缓存结果的专用样本库后，才重新评估 `.xlsm` 只读入口；任何版本都不会
执行宏代码。
