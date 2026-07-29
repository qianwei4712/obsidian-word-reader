import type { WordReaderLanguage } from "../i18n";

export interface XlsxReaderText {
  displayName: string;
  commands: {
    focusSearch: string;
    goToCell: string;
    copyValues: string;
    copyFormulas: string;
    copyMarkdown: string;
    createSummaryNote: string;
  };
  toolbar: {
    reload: string;
    searchPlaceholder: string;
    searchWorkbook: string;
    previousResult: string;
    nextResult: string;
    zoomPercentage: string;
    fitWidth: string;
    copyValues: string;
    copyFormulas: string;
    copyMarkdown: string;
    createSummaryNote: string;
    openExternally: string;
  };
  labels: {
    formula: string;
    cachedValue: string;
    openLink: string;
    sheetTabs: string;
    row: (row: number) => string;
    column: (column: string) => string;
    selectedRange: (range: string) => string;
    formulaSafety: string;
    nameBox: string;
    hiddenSheets: (count: number) => string;
    hiddenSheet: string;
    veryHiddenSheet: string;
  };
  status: {
    reading: (fileName: string) => string;
    parsingSheet: (sheetName: string) => string;
    ready: (
      fileName: string,
      sheetName: string,
      rows: number,
      columns: number,
    ) => string;
    searching: (
      sheetName: string,
      current: number,
      total: number,
    ) => string;
    searchResults: (count: number) => string;
    creatingSummary: (
      sheetName: string,
      current: number,
      total: number,
    ) => string;
    legacy: (fileName: string) => string;
  };
  notices: {
    copiedValues: string;
    copiedFormulas: string;
    copiedMarkdown: string;
    copyFailed: (message: string) => string;
    selectionTooLarge: (count: number, limit: number) => string;
    invalidReference: (value: string) => string;
    openedExistingSummaryNote: string;
    createdSummaryNote: string;
    summaryFailed: (message: string) => string;
    externalDesktopOnly: string;
    externalLocalVaultOnly: string;
    externalFailed: (message: string) => string;
    linkConfirmation: (target: string) => string;
    blockedLink: string;
  };
  summaryNote: {
    sourceLabel: string;
    currentSheetLabel: string;
    selectedRangeLabel: string;
    workbookOverviewHeading: string;
    namedRangesHeading: string;
    selectedRangeHeading: string;
    worksheetPreviewsHeading: string;
    keyFindingsHeading: string;
    followUpsHeading: string;
    cellHeader: string;
    displayedValueHeader: string;
    noNamedRanges: string;
    noPreview: string;
    visibility: {
      visible: string;
      hidden: string;
      veryHidden: string;
    };
    sheetSummary: (
      name: string,
      visibility: string,
      rows: number,
      columns: number,
      populatedCells: number,
    ) => string;
  };
  legacy: {
    title: string;
    body: string;
    convert: string;
  };
  errors: {
    formatMismatchTitle: string;
    formatMismatchBody: string;
    encryptedTitle: string;
    encryptedBody: string;
    limitTitle: string;
    limitBody: string;
    damagedTitle: string;
    damagedBody: string;
    unsupportedTitle: string;
    unsupportedBody: string;
    unknownTitle: string;
    unknownBody: string;
    tryAgain: string;
    openExternally: string;
    details: string;
  };
}

export function getXlsxReaderText(
  language: WordReaderLanguage,
): XlsxReaderText {
  return XLSX_READER_TEXT[language] ?? XLSX_READER_TEXT["zh-CN"];
}

const XLSX_READER_TEXT: Record<WordReaderLanguage, XlsxReaderText> = {
  en: {
    displayName: "Spreadsheet reader",
    commands: {
      focusSearch: "Search current workbook",
      goToCell: "Go to spreadsheet cell or named range",
      copyValues: "Copy selected spreadsheet values as TSV",
      copyFormulas: "Copy selected spreadsheet formulas",
      copyMarkdown: "Copy selected spreadsheet range as Markdown",
      createSummaryNote: "Create spreadsheet summary note",
    },
    toolbar: {
      reload: "Reload",
      searchPlaceholder: "Search workbook",
      searchWorkbook: "Search all worksheets",
      previousResult: "Previous result",
      nextResult: "Next result",
      zoomPercentage: "Zoom percentage",
      fitWidth: "Fit worksheet width",
      copyValues: "Copy displayed values as TSV",
      copyFormulas: "Copy formulas",
      copyMarkdown: "Copy as Markdown table",
      createSummaryNote: "Create workbook summary note",
      openExternally: "Open externally",
    },
    labels: {
      formula: "Formula",
      cachedValue: "Cached value",
      openLink: "Open link",
      sheetTabs: "Workbook sheets",
      row: (row) => `Row ${row}`,
      column: (column) => `Column ${column}`,
      selectedRange: (range) => `Selected ${range}`,
      formulaSafety:
        "Formulas are never recalculated; only stored text and cached values are shown.",
      nameBox: "Cell, range, or named range",
      hiddenSheets: (count) =>
        count === 1 ? "1 hidden sheet" : `${count} hidden sheets`,
      hiddenSheet: "Hidden worksheet",
      veryHiddenSheet: "Very hidden worksheet",
    },
    status: {
      reading: (fileName) => `Reading ${fileName}…`,
      parsingSheet: (sheetName) => `Loading worksheet ${sheetName}…`,
      ready: (fileName, sheetName, rows, columns) =>
        `${fileName} · ${sheetName} · ${rows.toLocaleString()} rows × ${columns.toLocaleString()} columns`,
      searching: (sheetName, current, total) =>
        `Searching ${sheetName} (${current}/${total})…`,
      searchResults: (count) =>
        count === 1 ? "1 result" : `${count.toLocaleString()} results`,
      creatingSummary: (sheetName, current, total) =>
        `Summarizing ${sheetName} (${current}/${total})…`,
      legacy: (fileName) => `${fileName} requires conversion to .xlsx`,
    },
    notices: {
      copiedValues: "Copied displayed cell values as TSV",
      copiedFormulas: "Copied formulas and displayed values",
      copiedMarkdown: "Copied selection as a Markdown table",
      copyFailed: (message) => `Could not copy cells: ${message}`,
      selectionTooLarge: (count, limit) =>
        `The selection contains ${count.toLocaleString()} cells. Copy at most ${limit.toLocaleString()} cells at once.`,
      invalidReference: (value) =>
        `Could not find cell, range, or named range: ${value}`,
      openedExistingSummaryNote: "Opened existing spreadsheet summary note",
      createdSummaryNote: "Created spreadsheet summary note",
      summaryFailed: (message) =>
        `Could not create spreadsheet summary: ${message}`,
      externalDesktopOnly:
        "External opening is only available in Obsidian Desktop",
      externalLocalVaultOnly: "External opening requires a local desktop vault",
      externalFailed: (message) => `Could not open workbook: ${message}`,
      linkConfirmation: (target) => `Open this workbook link?\n\n${target}`,
      blockedLink: "Only safe web, email, and in-workbook links can be opened",
    },
    summaryNote: {
      sourceLabel: "Source",
      currentSheetLabel: "Current worksheet",
      selectedRangeLabel: "Selected range",
      workbookOverviewHeading: "Workbook overview",
      namedRangesHeading: "Named ranges",
      selectedRangeHeading: "Selected range snapshot",
      worksheetPreviewsHeading: "Worksheet previews",
      keyFindingsHeading: "Key findings",
      followUpsHeading: "Follow-ups",
      cellHeader: "Cell",
      displayedValueHeader: "Displayed value",
      noNamedRanges: "No supported named ranges were found.",
      noPreview: "No populated cells are available for this preview.",
      visibility: {
        visible: "visible",
        hidden: "hidden",
        veryHidden: "very hidden",
      },
      sheetSummary: (name, visibility, rows, columns, populatedCells) =>
        `${name} (${visibility}) — ${rows.toLocaleString()} rows × ${columns.toLocaleString()} columns; ${populatedCells.toLocaleString()} populated cells`,
    },
    legacy: {
      title: "Legacy Excel workbook",
      body: "Binary .xls files cannot be rendered safely inside Obsidian.",
      convert:
        "Open the file externally and save a copy as .xlsx to use the built-in read-only grid.",
    },
    errors: {
      formatMismatchTitle: "Not a valid XLSX workbook",
      formatMismatchBody:
        "The file is not a supported ZIP-based Excel workbook.",
      encryptedTitle: "Encrypted workbook",
      encryptedBody:
        "Encrypted or password-protected workbooks cannot be previewed.",
      limitTitle: "Workbook exceeds safe limits",
      limitBody:
        "The workbook is outside the safe package or memory envelope.",
      damagedTitle: "Damaged workbook",
      damagedBody:
        "The workbook package or worksheet XML could not be read safely.",
      unsupportedTitle: "Unsupported workbook content",
      unsupportedBody:
        "This workbook contains structures that the read-only preview does not support.",
      unknownTitle: "Could not open workbook",
      unknownBody: "An unexpected error stopped the spreadsheet preview.",
      tryAgain: "Reload the workbook after verifying the source file.",
      openExternally: "Open the original file in a trusted spreadsheet app.",
      details: "Diagnostic details",
    },
  },
  "zh-CN": {
    displayName: "电子表格阅读器",
    commands: {
      focusSearch: "搜索当前工作簿",
      goToCell: "跳转到单元格或命名区域",
      copyValues: "复制选中区域显示值为 TSV",
      copyFormulas: "复制选中区域公式",
      copyMarkdown: "复制选区为 Markdown 表格",
      createSummaryNote: "创建工作簿摘要笔记",
    },
    toolbar: {
      reload: "重新加载",
      searchPlaceholder: "搜索整个工作簿",
      searchWorkbook: "搜索全部工作表",
      previousResult: "上一个结果",
      nextResult: "下一个结果",
      zoomPercentage: "缩放百分比",
      fitWidth: "适配工作表宽度",
      copyValues: "复制显示值为 TSV",
      copyFormulas: "复制公式",
      copyMarkdown: "复制为 Markdown 表格",
      createSummaryNote: "创建工作簿摘要笔记",
      openExternally: "外部打开",
    },
    labels: {
      formula: "公式",
      cachedValue: "缓存结果",
      openLink: "打开链接",
      sheetTabs: "工作簿工作表",
      row: (row) => `第 ${row} 行`,
      column: (column) => `第 ${column} 列`,
      selectedRange: (range) => `已选择 ${range}`,
      formulaSafety: "不会重新计算公式，只显示文件中保存的公式文本和缓存结果。",
      nameBox: "单元格、区域或命名区域",
      hiddenSheets: (count) => `${count} 个隐藏工作表`,
      hiddenSheet: "隐藏工作表",
      veryHiddenSheet: "深度隐藏工作表",
    },
    status: {
      reading: (fileName) => `正在读取 ${fileName}…`,
      parsingSheet: (sheetName) => `正在加载工作表 ${sheetName}…`,
      ready: (fileName, sheetName, rows, columns) =>
        `${fileName} · ${sheetName} · ${rows.toLocaleString()} 行 × ${columns.toLocaleString()} 列`,
      searching: (sheetName, current, total) =>
        `正在搜索 ${sheetName}（${current}/${total}）…`,
      searchResults: (count) => `${count.toLocaleString()} 个结果`,
      creatingSummary: (sheetName, current, total) =>
        `正在汇总 ${sheetName}（${current}/${total}）…`,
      legacy: (fileName) => `${fileName} 需要转换为 .xlsx`,
    },
    notices: {
      copiedValues: "已将单元格显示值复制为 TSV",
      copiedFormulas: "已复制公式和非公式单元格的显示值",
      copiedMarkdown: "已将选区复制为 Markdown 表格",
      copyFailed: (message) => `无法复制单元格：${message}`,
      selectionTooLarge: (count, limit) =>
        `选区包含 ${count.toLocaleString()} 个单元格，请一次最多复制 ${limit.toLocaleString()} 个。`,
      invalidReference: (value) =>
        `未找到单元格、区域或命名区域：${value}`,
      openedExistingSummaryNote: "已打开现有电子表格摘要笔记",
      createdSummaryNote: "已创建电子表格摘要笔记",
      summaryFailed: (message) => `无法创建电子表格摘要：${message}`,
      externalDesktopOnly: "外部打开仅支持 Obsidian 桌面端",
      externalLocalVaultOnly: "外部打开需要本地桌面仓库",
      externalFailed: (message) => `无法打开工作簿：${message}`,
      linkConfirmation: (target) => `是否打开此工作簿链接？\n\n${target}`,
      blockedLink: "只能打开安全的网页、邮件和工作簿内部链接",
    },
    summaryNote: {
      sourceLabel: "来源",
      currentSheetLabel: "当前工作表",
      selectedRangeLabel: "选中区域",
      workbookOverviewHeading: "工作簿概览",
      namedRangesHeading: "命名区域",
      selectedRangeHeading: "选区快照",
      worksheetPreviewsHeading: "工作表预览",
      keyFindingsHeading: "关键发现",
      followUpsHeading: "后续事项",
      cellHeader: "单元格",
      displayedValueHeader: "显示值",
      noNamedRanges: "未发现受支持的命名区域。",
      noPreview: "没有可用于预览的非空单元格。",
      visibility: {
        visible: "可见",
        hidden: "隐藏",
        veryHidden: "深度隐藏",
      },
      sheetSummary: (name, visibility, rows, columns, populatedCells) =>
        `${name}（${visibility}）— ${rows.toLocaleString()} 行 × ${columns.toLocaleString()} 列；${populatedCells.toLocaleString()} 个非空单元格`,
    },
    legacy: {
      title: "旧版 Excel 工作簿",
      body: "无法在 Obsidian 内安全渲染二进制 .xls 文件。",
      convert: "请在外部打开并另存为 .xlsx，再使用内置只读网格阅读。",
    },
    errors: {
      formatMismatchTitle: "不是有效的 XLSX 工作簿",
      formatMismatchBody: "该文件不是受支持的 ZIP 格式 Excel 工作簿。",
      encryptedTitle: "工作簿已加密",
      encryptedBody: "无法预览加密或受密码保护的工作簿。",
      limitTitle: "工作簿超出安全限制",
      limitBody: "该工作簿超出安全包或内存范围。",
      damagedTitle: "工作簿已损坏",
      damagedBody: "无法安全读取工作簿包或工作表 XML。",
      unsupportedTitle: "不支持的工作簿内容",
      unsupportedBody: "该工作簿包含只读预览尚不支持的结构。",
      unknownTitle: "无法打开工作簿",
      unknownBody: "电子表格预览遇到意外错误。",
      tryAgain: "确认源文件正常后重新加载工作簿。",
      openExternally: "在可信的电子表格应用中打开原文件。",
      details: "诊断详情",
    },
  },
};
