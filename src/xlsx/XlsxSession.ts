import {
  FileView,
  Notice,
  TFile,
  WorkspaceLeaf,
  setIcon,
} from "obsidian";

import { openExternalFile } from "../commands/openExternal";
import type WordReaderPlugin from "../main";
import type { ReaderCapability } from "../reader/capabilities";
import { ReaderLifecycle } from "../reader/lifecycle";
import type { ReaderViewState } from "../reader/readingState";
import type { ReaderSession } from "../reader/session";
import { OfficeReaderShell } from "../reader/shell";
import {
  ReaderStatusController,
  type ReaderStatus,
} from "../reader/status";
import { normalizeZoom, preserveZoomAnchor } from "../reader/zoom";
import { columnIndexToName, parseCellReference } from "./xlsxReferences";
import {
  MAX_XLSX_COPY_CELLS,
  normalizeXlsxSelection,
  xlsxSelectionContains,
  xlsxSelectionToTsv,
  XlsxSelectionTooLargeError,
  type XlsxCellPosition,
  type XlsxCopyMode,
} from "./xlsxSelection";
import { xlsxCellStyleToCss } from "./xlsxStyleCss";
import {
  XLSX_ADAPTER,
  XLSX_VIEW_TYPE,
  type XlsxAdapter,
} from "./XlsxAdapter";
import { classifyXlsxError } from "./xlsxErrors";
import {
  isSafeXlsxExternalTarget,
  parseXlsxWorkbookLocation,
} from "./xlsxHyperlinks";
import {
  getXlsxReaderText,
  type XlsxReaderText,
} from "./xlsxI18n";
import { XlsxPackage } from "./xlsxPackage";
import {
  searchXlsxCells,
  XlsxSearchCancelledError,
} from "./xlsxSearch";
import type {
  XlsxCell,
  XlsxHyperlink,
  XlsxMergeRange,
} from "./xlsxTypes";
import { XlsxVirtualGrid } from "./xlsxVirtualGrid";
import {
  XlsxWorksheet,
  XlsxWorksheetCancelledError,
} from "./xlsxWorksheet";

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4;
const ZOOM_STEP = 0.05;
const SEARCH_DEBOUNCE_MS = 160;
let xlsxSessionSequence = 0;

interface CellCoordinate {
  row: number;
  column: number;
}

export class XlsxSession extends FileView implements ReaderSession {
  readonly adapter: XlsxAdapter = XLSX_ADAPTER;
  readonly capabilities = this.adapter.capabilities;

  private readonly plugin: WordReaderPlugin;
  private readonly shell: OfficeReaderShell;
  private readonly cellIdPrefix = `xlsx-reader-${++xlsxSessionSequence}`;
  private rootEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private bodyEl: HTMLElement | null = null;
  private formulaNameEl: HTMLElement | null = null;
  private formulaValueEl: HTMLElement | null = null;
  private formulaCachedEl: HTMLElement | null = null;
  private formulaSafetyEl: HTMLElement | null = null;
  private linkButtonEl: HTMLButtonElement | null = null;
  private gridShellEl: HTMLElement | null = null;
  private viewportEl: HTMLElement | null = null;
  private surfaceEl: HTMLElement | null = null;
  private cellsEl: HTMLElement | null = null;
  private rowHeadersEl: HTMLElement | null = null;
  private columnHeadersEl: HTMLElement | null = null;
  private cornerEl: HTMLElement | null = null;
  private sheetTabsEl: HTMLElement | null = null;
  private searchInputEl: HTMLInputElement | null = null;
  private searchCountEl: HTMLElement | null = null;
  private zoomInputEl: HTMLInputElement | null = null;
  private fitButtonEl: HTMLButtonElement | null = null;
  private gridMessageEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private workbook: XlsxPackage | null = null;
  private worksheet: XlsxWorksheet | null = null;
  private virtualGrid: XlsxVirtualGrid | null = null;
  private activeSheetIndex = 0;
  private activeFilePath: string | null = null;
  private zoom = 1;
  private fitWidth = true;
  private pendingScrollPosition: { left: number; top: number } | null = null;
  private selectionAnchor: XlsxCellPosition = { row: 0, column: 0 };
  private selectionFocus: XlsxCellPosition = { row: 0, column: 0 };
  private draggingSelection = false;
  private searchQuery = "";
  private searchResults: XlsxCellPosition[] = [];
  private searchResultKeys = new Set<string>();
  private currentSearchResult = -1;
  private searchTimer: number | null = null;
  private renderFrameId: number | null = null;
  private stateSaveTimer: number | null = null;
  private readonly loadLifecycle = new ReaderLifecycle();
  private readonly sheetLifecycle = new ReaderLifecycle();
  private readonly searchLifecycle = new ReaderLifecycle();
  private readonly readerStatus = new ReaderStatusController((status) => {
    this.renderStatus(status);
  });
  private readonly handlePointerUp = (): void => {
    this.draggingSelection = false;
  };

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.shell = new OfficeReaderShell(this.contentEl, this.capabilities);
    this.zoom = plugin.settings.common.defaultZoomPercent / 100;
    this.fitWidth = plugin.settings.xlsx.defaultFitWidth;
  }

  getViewType(): string {
    return XLSX_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file?.name ?? this.text.displayName;
  }

  getIcon(): string {
    return "table-2";
  }

  private get text(): XlsxReaderText {
    return getXlsxReaderText(this.plugin.settings.common.language);
  }

  async onOpen(): Promise<void> {
    this.buildLayout();
    this.contentEl.doc.addEventListener("pointerup", this.handlePointerUp);
    this.resizeObserver = new ResizeObserver(() => {
      this.applyScale();
      this.scheduleGridRender();
    });
    this.resizeObserver.observe(this.contentEl);
  }

  async onClose(): Promise<void> {
    this.contentEl.doc.removeEventListener("pointerup", this.handlePointerUp);
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    this.cancelScheduledWork();
    this.loadLifecycle.cancel();
    this.sheetLifecycle.cancel();
    this.searchLifecycle.cancel();
    this.releaseWorkbook();
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.restoreReadingState(file);
    await this.loadWorkbook(file);
  }

  async onUnloadFile(): Promise<void> {
    this.saveReadingState();
    await this.plugin.flushData();
    this.activeFilePath = null;
    this.loadLifecycle.cancel();
    this.sheetLifecycle.cancel();
    this.searchLifecycle.cancel();
    this.cancelScheduledWork();
    this.releaseWorkbook();
    this.setStatus("");
  }

  async reload(): Promise<void> {
    if (this.file) {
      await this.loadWorkbook(this.file);
    }
  }

  async openExternal(): Promise<void> {
    if (!this.file) {
      return;
    }
    try {
      await openExternalFile(this.app, this.file, {
        desktopOnly: this.text.notices.externalDesktopOnly,
        localVaultOnly: this.text.notices.externalLocalVaultOnly,
      });
    } catch (error) {
      new Notice(this.text.notices.externalFailed(getErrorMessage(error)));
    }
  }

  async copyText(): Promise<void> {
    await this.copySelection("display");
  }

  async copyFormulas(): Promise<void> {
    await this.copySelection("formula");
  }

  focusSearch(): void {
    this.searchInputEl?.focus();
    this.searchInputEl?.select();
  }

  refreshInterfaceLanguage(): void {
    const viewport = this.viewportEl;
    const scrollPosition = viewport
      ? { left: viewport.scrollLeft, top: viewport.scrollTop }
      : null;
    this.buildLayout();
    this.pendingScrollPosition = scrollPosition;
    if (this.file?.extension.toLowerCase() === "xls") {
      this.showLegacyWorkbook();
      return;
    }
    if (this.workbook && this.worksheet) {
      this.renderSheetTabs();
      this.updateFormulaBar();
      this.applyScale();
      this.restoreScrollPosition();
      this.scheduleGridRender();
      this.setReadyStatus();
    }
    this.readerStatus.refresh();
  }

  private buildLayout(): void {
    this.cancelScheduledWork();
    const text = this.text;
    const layout = this.shell.build({
      rootClasses: "xlsx-reader-root",
      toolbarClasses: "xlsx-reader-toolbar",
      statusClasses: "xlsx-reader-status",
      bodyClasses: "xlsx-reader-body",
    });
    this.rootEl = layout.rootEl;
    this.statusEl = layout.statusEl;
    this.bodyEl = layout.bodyEl;
    this.rootEl.tabIndex = 0;
    this.rootEl.addEventListener("keydown", (event) => {
      this.handleKeyDown(event);
    });

    this.createToolbarButton(
      layout.toolbarEl,
      "reload",
      "refresh-cw",
      text.toolbar.reload,
      () => {
        void this.reload();
      },
    );

    const searchEl = layout.toolbarEl.createDiv({
      cls: "xlsx-reader-search-control",
    });
    this.searchInputEl = searchEl.createEl("input", {
      cls: "office-reader-search xlsx-reader-search",
      attr: {
        type: "search",
        placeholder: text.toolbar.searchPlaceholder,
        "aria-label": text.toolbar.searchWorkbook,
      },
    });
    this.searchInputEl.value = this.searchQuery;
    this.searchInputEl.addEventListener("input", () => {
      this.searchQuery = this.searchInputEl?.value ?? "";
      this.scheduleSearch();
    });
    this.searchInputEl.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.goToSearchResult(event.shiftKey ? -1 : 1);
      }
    });
    this.createIconButton(
      searchEl,
      "chevron-up",
      text.toolbar.previousResult,
      () => {
        this.goToSearchResult(-1);
      },
    );
    this.createIconButton(
      searchEl,
      "chevron-down",
      text.toolbar.nextResult,
      () => {
        this.goToSearchResult(1);
      },
    );
    this.searchCountEl = searchEl.createSpan({
      cls: "office-reader-search-count xlsx-reader-search-count",
    });

    this.zoomInputEl = layout.toolbarEl.createEl("input", {
      cls: "office-reader-zoom xlsx-reader-zoom",
      attr: {
        type: "number",
        min: String(MIN_ZOOM * 100),
        max: String(MAX_ZOOM * 100),
        step: String(ZOOM_STEP * 100),
        "aria-label": text.toolbar.zoomPercentage,
        title: text.toolbar.zoomPercentage,
      },
    });
    this.zoomInputEl.addEventListener("change", () => {
      this.setZoom(Number(this.zoomInputEl?.value) / 100);
    });
    this.fitButtonEl = this.createToolbarButton(
      layout.toolbarEl,
      "fit",
      "move-horizontal",
      text.toolbar.fitWidth,
      () => {
        this.fitWidth = !this.fitWidth;
        if (this.fitWidth && this.viewportEl) {
          this.viewportEl.scrollLeft = 0;
        }
        this.applyScale();
        this.saveReadingState();
      },
    );
    this.createToolbarButton(
      layout.toolbarEl,
      "copyText",
      "copy",
      text.toolbar.copyValues,
      () => {
        void this.copySelection("display");
      },
    );
    this.createToolbarButton(
      layout.toolbarEl,
      "copyText",
      "function-square",
      text.toolbar.copyFormulas,
      () => {
        void this.copySelection("formula");
      },
    );
    this.createToolbarButton(
      layout.toolbarEl,
      "openExternal",
      "external-link",
      text.toolbar.openExternally,
      () => {
        void this.openExternal();
      },
    );

    const formulaBarEl = this.bodyEl.createDiv({
      cls: "xlsx-reader-formula-bar",
    });
    this.formulaNameEl = formulaBarEl.createDiv({
      cls: "xlsx-reader-name-box",
    });
    this.formulaValueEl = formulaBarEl.createDiv({
      cls: "xlsx-reader-formula-value",
    });
    this.formulaCachedEl = formulaBarEl.createDiv({
      cls: "xlsx-reader-cached-value",
    });
    this.formulaSafetyEl = formulaBarEl.createDiv({
      cls: "xlsx-reader-formula-safety",
      text: text.labels.formulaSafety,
    });
    this.linkButtonEl = this.createIconButton(
      formulaBarEl,
      "external-link",
      text.labels.openLink,
      () => {
        void this.openSelectedHyperlink();
      },
    );

    this.gridShellEl = this.bodyEl.createDiv({
      cls: "xlsx-reader-grid-shell",
    });
    this.cornerEl = this.gridShellEl.createDiv({
      cls: "xlsx-reader-corner",
    });
    this.columnHeadersEl = this.gridShellEl.createDiv({
      cls: "xlsx-reader-column-headers",
      attr: { "aria-hidden": "true" },
    });
    this.rowHeadersEl = this.gridShellEl.createDiv({
      cls: "xlsx-reader-row-headers",
      attr: { "aria-hidden": "true" },
    });
    this.viewportEl = this.gridShellEl.createDiv({
      cls: "xlsx-reader-viewport",
      attr: {
        role: "grid",
        "aria-label": text.displayName,
      },
    });
    this.viewportEl.addEventListener("scroll", () => {
      this.scheduleGridRender();
      this.scheduleStateSave();
    });
    this.viewportEl.addEventListener("wheel", (event) => {
      this.handleWheelZoom(event);
    });
    this.surfaceEl = this.viewportEl.createDiv({
      cls: "xlsx-reader-surface",
    });
    this.cellsEl = this.surfaceEl.createDiv({
      cls: "xlsx-reader-cells",
    });
    this.sheetTabsEl = this.bodyEl.createDiv({
      cls: "xlsx-reader-sheet-tabs",
      attr: {
        role: "tablist",
        "aria-label": text.labels.sheetTabs,
      },
    });

    this.updateFormulaBar();
    this.updateSearchCount();
    this.applyScale();
    this.readerStatus.refresh();
  }

  private async loadWorkbook(file: TFile): Promise<void> {
    const token = this.loadLifecycle.begin();
    this.sheetLifecycle.cancel();
    this.searchLifecycle.cancel();
    this.releaseWorkbook();
    this.clearGridMessage();
    this.searchResults = [];
    this.searchResultKeys.clear();
    this.currentSearchResult = -1;
    this.updateSearchCount();

    if (file.extension.toLowerCase() === "xls") {
      this.showLegacyWorkbook();
      return;
    }

    this.setStatus(this.text.status.reading(file.name), false, true);
    try {
      const workbook = await this.adapter.open(this.app, file);
      if (!this.loadLifecycle.isCurrent(token)) {
        workbook.clearCaches();
        return;
      }
      this.workbook = workbook;
      this.activeSheetIndex = this.normalizeSheetIndex(
        this.activeSheetIndex,
        workbook,
      );
      this.renderSheetTabs();
      await this.loadWorksheet(this.activeSheetIndex, true);
    } catch (error) {
      if (this.loadLifecycle.isCurrent(token)) {
        this.releaseWorkbook();
        this.showError(error);
      }
    }
  }

  private async loadWorksheet(
    index: number,
    restoreScroll: boolean,
  ): Promise<void> {
    const workbook = this.workbook;
    const descriptor = workbook?.sheets[index];
    if (!workbook || !descriptor) {
      return;
    }

    const token = this.sheetLifecycle.begin();
    this.searchLifecycle.cancel();
    this.clearWorksheet();
    this.activeSheetIndex = index;
    this.renderSheetTabs();
    this.setStatus(
      this.text.status.parsingSheet(descriptor.name),
      false,
      true,
    );

    try {
      const worksheet = await workbook.getWorksheet(index, {
        isCancelled: () =>
          !this.sheetLifecycle.isCurrent(token) ||
          this.workbook !== workbook,
      });
      if (
        !this.sheetLifecycle.isCurrent(token) ||
        this.workbook !== workbook
      ) {
        return;
      }
      this.worksheet = worksheet;
      this.virtualGrid = new XlsxVirtualGrid(worksheet);
      this.selectionAnchor = { row: 0, column: 0 };
      this.selectionFocus = { row: 0, column: 0 };
      this.updateGridAria();
      this.updateFormulaBar();
      this.applyScale();
      if (restoreScroll) {
        this.restoreScrollPosition();
      } else if (this.viewportEl) {
        this.viewportEl.scrollLeft = 0;
        this.viewportEl.scrollTop = 0;
      }
      this.scheduleGridRender();
      this.setReadyStatus();
      this.saveReadingState();
      void this.runSearch();
    } catch (error) {
      if (error instanceof XlsxWorksheetCancelledError) {
        return;
      }
      if (
        this.sheetLifecycle.isCurrent(token) &&
        this.workbook === workbook
      ) {
        this.showError(error);
      }
    }
  }

  private clearWorksheet(): void {
    this.worksheet = null;
    this.virtualGrid = null;
    this.cellsEl?.empty();
    this.rowHeadersEl?.empty();
    this.columnHeadersEl?.empty();
    if (this.surfaceEl) {
      this.surfaceEl.setCssProps({
        width: "0",
        height: "0",
      });
    }
  }

  private releaseWorkbook(): void {
    this.workbook?.clearCaches();
    this.workbook = null;
    this.clearWorksheet();
    this.sheetTabsEl?.empty();
  }

  private renderSheetTabs(): void {
    const tabsEl = this.sheetTabsEl;
    const workbook = this.workbook;
    if (!tabsEl || !workbook) {
      return;
    }
    tabsEl.empty();
    const visible = workbook.sheets
      .map((sheet, index) => ({ sheet, index }))
      .filter(({ sheet }) => sheet.state === "visible");
    if (!visible.some(({ index }) => index === this.activeSheetIndex)) {
      const active = workbook.sheets[this.activeSheetIndex];
      if (active) {
        visible.push({ sheet: active, index: this.activeSheetIndex });
      }
    }
    for (const { sheet, index } of visible) {
      const buttonEl = tabsEl.createEl("button", {
        cls: "xlsx-reader-sheet-tab",
        text: sheet.name,
        attr: {
          type: "button",
          role: "tab",
          "aria-selected": String(index === this.activeSheetIndex),
          title: sheet.name,
        },
      });
      buttonEl.toggleClass("is-active", index === this.activeSheetIndex);
      buttonEl.toggleClass("is-hidden-sheet", sheet.state !== "visible");
      buttonEl.addEventListener("click", () => {
        if (index !== this.activeSheetIndex) {
          this.pendingScrollPosition = { left: 0, top: 0 };
          void this.loadWorksheet(index, false);
        }
      });
    }
  }

  private scheduleGridRender(): void {
    if (this.renderFrameId !== null) {
      return;
    }
    this.renderFrameId = this.contentEl.win.requestAnimationFrame(() => {
      this.renderFrameId = null;
      this.renderGrid();
    });
  }

  private renderGrid(): void {
    const worksheet = this.worksheet;
    const grid = this.virtualGrid;
    const viewportEl = this.viewportEl;
    const surfaceEl = this.surfaceEl;
    const cellsEl = this.cellsEl;
    const rowHeadersEl = this.rowHeadersEl;
    const columnHeadersEl = this.columnHeadersEl;
    if (
      !worksheet ||
      !grid ||
      !viewportEl ||
      !surfaceEl ||
      !cellsEl ||
      !rowHeadersEl ||
      !columnHeadersEl
    ) {
      return;
    }

    const scale = this.getEffectiveScale();
    surfaceEl.style.width = `${grid.totalWidth * scale}px`;
    surfaceEl.style.height = `${grid.totalHeight * scale}px`;
    const virtualWindow = grid.calculate({
      scrollTop: viewportEl.scrollTop / scale,
      scrollLeft: viewportEl.scrollLeft / scale,
      width: viewportEl.clientWidth / scale,
      height: viewportEl.clientHeight / scale,
    });
    const baseRows = range(virtualWindow.startRow, virtualWindow.endRow);
    const baseColumns = range(
      virtualWindow.startColumn,
      virtualWindow.endColumn,
    );
    const frozenRows = this.getMountedFrozenRows(
      worksheet,
      grid,
      viewportEl.clientHeight / scale,
    );
    const frozenColumns = this.getMountedFrozenColumns(
      worksheet,
      grid,
      viewportEl.clientWidth / scale,
    );
    const rows = uniqueSorted([...frozenRows, ...baseRows]);
    const columns = uniqueSorted([...frozenColumns, ...baseColumns]);

    cellsEl.empty();
    rowHeadersEl.empty();
    columnHeadersEl.empty();
    this.renderHeaders(rows, columns, frozenRows, frozenColumns, scale);

    const visibleMerges = worksheet.merges.filter((merge) =>
      mergeIntersectsAxes(merge, rows, columns),
    );
    const mounted = new Set<string>();
    const mountedMerges = new Set<string>();
    const prioritizedCoordinates = [
      ...cartesian(frozenRows, columns),
      ...cartesian(rows, frozenColumns),
      ...cartesian(baseRows, baseColumns),
    ];
    for (const coordinate of prioritizedCoordinates) {
      if (mounted.size >= grid.mountBudget) {
        break;
      }
      const key = cellKey(coordinate.row, coordinate.column);
      if (mounted.has(key)) {
        continue;
      }
      const merge = visibleMerges.find((candidate) =>
        mergeContains(candidate, coordinate.row, coordinate.column),
      );
      if (merge) {
        if (mountedMerges.has(merge.ref)) {
          continue;
        }
        mountedMerges.add(merge.ref);
        this.renderCell(
          merge.startRow,
          merge.startColumn,
          merge,
          scale,
        );
        mounted.add(cellKey(merge.startRow, merge.startColumn));
        continue;
      }
      this.renderCell(coordinate.row, coordinate.column, null, scale);
      mounted.add(key);
    }
    this.updateFormulaBar();
  }

  private renderHeaders(
    rows: readonly number[],
    columns: readonly number[],
    frozenRows: readonly number[],
    frozenColumns: readonly number[],
    scale: number,
  ): void {
    const grid = this.virtualGrid;
    const viewportEl = this.viewportEl;
    const rowHeadersEl = this.rowHeadersEl;
    const columnHeadersEl = this.columnHeadersEl;
    if (!grid || !viewportEl || !rowHeadersEl || !columnHeadersEl) {
      return;
    }
    const frozenRowSet = new Set(frozenRows);
    const frozenColumnSet = new Set(frozenColumns);

    for (const column of columns) {
      const headerEl = columnHeadersEl.createDiv({
        cls: "xlsx-reader-column-header",
        text: columnIndexToName(column),
        attr: {
          title: this.text.labels.column(columnIndexToName(column)),
        },
      });
      const frozen = frozenColumnSet.has(column);
      headerEl.toggleClass("is-frozen", frozen);
      headerEl.style.left = `${
        grid.columnOffset(column) * scale -
        (frozen ? 0 : viewportEl.scrollLeft)
      }px`;
      headerEl.style.width = `${grid.columnSize(column) * scale}px`;
    }
    for (const row of rows) {
      const headerEl = rowHeadersEl.createDiv({
        cls: "xlsx-reader-row-header",
        text: String(row + 1),
        attr: {
          title: this.text.labels.row(row + 1),
        },
      });
      const frozen = frozenRowSet.has(row);
      headerEl.toggleClass("is-frozen", frozen);
      headerEl.style.top = `${
        grid.rowOffset(row) * scale -
        (frozen ? 0 : viewportEl.scrollTop)
      }px`;
      headerEl.style.height = `${grid.rowSize(row) * scale}px`;
    }
  }

  private renderCell(
    row: number,
    column: number,
    merge: XlsxMergeRange | null,
    scale: number,
  ): void {
    const worksheet = this.worksheet;
    const grid = this.virtualGrid;
    const viewportEl = this.viewportEl;
    const cellsEl = this.cellsEl;
    if (!worksheet || !grid || !viewportEl || !cellsEl) {
      return;
    }
    const cell = worksheet.getCell(row, column);
    const endRow = merge?.endRow ?? row;
    const endColumn = merge?.endColumn ?? column;
    const frozenRows = worksheet.frozenPane?.rows ?? 0;
    const frozenColumns = worksheet.frozenPane?.columns ?? 0;
    const frozenRow = row < frozenRows;
    const frozenColumn = column < frozenColumns;
    const cellEl = cellsEl.createDiv({
      cls: "xlsx-reader-cell",
      attr: {
        role: "gridcell",
        "aria-rowindex": String(row + 1),
        "aria-colindex": String(column + 1),
        "data-row": String(row),
        "data-column": String(column),
        id: this.getCellElementId(row, column),
      },
    });
    cellEl.style.left = `${
      grid.columnOffset(column) * scale +
      (frozenColumn ? viewportEl.scrollLeft : 0)
    }px`;
    cellEl.style.top = `${
      grid.rowOffset(row) * scale +
      (frozenRow ? viewportEl.scrollTop : 0)
    }px`;
    cellEl.style.width = `${
      (grid.columnOffset(endColumn + 1) - grid.columnOffset(column)) * scale
    }px`;
    cellEl.style.height = `${
      (grid.rowOffset(endRow + 1) - grid.rowOffset(row)) * scale
    }px`;
    cellEl.style.zIndex = String(
      frozenRow && frozenColumn ? 4 : frozenRow || frozenColumn ? 3 : 1,
    );

    if (cell) {
      const css = xlsxCellStyleToCss(cell.style);
      const rotation = css.transform;
      delete css.transform;
      Object.assign(cellEl.style, css);
      const contentEl = cellEl.createSpan({
        cls: "xlsx-reader-cell-content",
        text: cell.displayValue,
      });
      if (rotation) {
        contentEl.style.transform = rotation;
      }
      cellEl.toggleClass("has-formula", Boolean(cell.formula));
      cellEl.toggleClass("has-hyperlink", Boolean(cell.hyperlink));
      cellEl.title = getCellTitle(cell);
    }
    const selection = this.selection;
    cellEl.toggleClass(
      "is-selected",
      xlsxSelectionContains(selection, row, column),
    );
    cellEl.toggleClass(
      "is-selection-focus",
      row === this.selectionFocus.row &&
        column === this.selectionFocus.column,
    );
    cellEl.toggleClass(
      "is-search-match",
      this.searchResultKeys.has(cellKey(row, column)),
    );
    cellEl.addEventListener("pointerdown", (event) => {
      this.handleCellPointerDown(event, row, column);
    });
    cellEl.addEventListener("pointerenter", () => {
      if (this.draggingSelection) {
        this.selectionFocus = this.clampCellPosition({ row, column });
        this.updateFormulaBar();
        this.scheduleGridRender();
      }
    });
    cellEl.addEventListener("dblclick", () => {
      if (cell?.hyperlink) {
        void this.openHyperlink(cell.hyperlink);
      }
    });
  }

  private getMountedFrozenRows(
    worksheet: XlsxWorksheet,
    grid: XlsxVirtualGrid,
    viewportHeight: number,
  ): number[] {
    const count = worksheet.frozenPane?.rows ?? 0;
    if (count === 0) {
      return [];
    }
    const fitCount = grid.rowIndexAtOffset(viewportHeight) + 1;
    return range(0, Math.min(count, fitCount, worksheet.rowCount));
  }

  private getMountedFrozenColumns(
    worksheet: XlsxWorksheet,
    grid: XlsxVirtualGrid,
    viewportWidth: number,
  ): number[] {
    const count = worksheet.frozenPane?.columns ?? 0;
    if (count === 0) {
      return [];
    }
    const fitCount = grid.columnIndexAtOffset(viewportWidth) + 1;
    return range(0, Math.min(count, fitCount, worksheet.columnCount));
  }

  private handleCellPointerDown(
    event: PointerEvent,
    row: number,
    column: number,
  ): void {
    if (event.button !== 0) {
      return;
    }
    event.preventDefault();
    const position = this.clampCellPosition({ row, column });
    if (!event.shiftKey) {
      this.selectionAnchor = position;
    }
    this.selectionFocus = position;
    this.draggingSelection = true;
    this.rootEl?.focus();
    this.updateFormulaBar();
    this.scheduleGridRender();
  }

  private handleKeyDown(event: KeyboardEvent): void {
    const target = event.target as Node | null;
    if (
      target?.instanceOf(HTMLElement) &&
      target.closest("input, textarea, select, button, a, [contenteditable='true']")
    ) {
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f") {
      event.preventDefault();
      this.focusSearch();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      event.preventDefault();
      void this.copySelection("display");
      return;
    }
    if (event.altKey || event.ctrlKey || event.metaKey) {
      return;
    }

    let rowDelta = 0;
    let columnDelta = 0;
    switch (event.key) {
      case "ArrowUp":
        rowDelta = -1;
        break;
      case "ArrowDown":
      case "Enter":
        rowDelta = 1;
        break;
      case "ArrowLeft":
        columnDelta = -1;
        break;
      case "ArrowRight":
        columnDelta = 1;
        break;
      case "PageUp":
        rowDelta = -this.visibleRowStep;
        break;
      case "PageDown":
        rowDelta = this.visibleRowStep;
        break;
      case "Home":
        columnDelta = -this.selectionFocus.column;
        break;
      case "End":
        columnDelta =
          (this.worksheet?.columnCount ?? 1) -
          this.selectionFocus.column -
          1;
        break;
      default:
        return;
    }
    event.preventDefault();
    const next = this.clampCellPosition({
      row: this.selectionFocus.row + rowDelta,
      column: this.selectionFocus.column + columnDelta,
    });
    if (!event.shiftKey) {
      this.selectionAnchor = next;
    }
    this.selectionFocus = next;
    this.scrollCellIntoView(next);
    this.updateFormulaBar();
    this.scheduleGridRender();
  }

  private get visibleRowStep(): number {
    const grid = this.virtualGrid;
    const viewportEl = this.viewportEl;
    const scale = this.getEffectiveScale();
    if (!grid || !viewportEl) {
      return 1;
    }
    const first = grid.rowIndexAtOffset(viewportEl.scrollTop / scale);
    const last = grid.rowIndexAtOffset(
      (viewportEl.scrollTop + viewportEl.clientHeight) / scale,
    );
    return Math.max(1, last - first);
  }

  private get selection() {
    return normalizeXlsxSelection(
      this.selectionAnchor,
      this.selectionFocus,
    );
  }

  private clampCellPosition(position: XlsxCellPosition): XlsxCellPosition {
    const worksheet = this.worksheet;
    return {
      row: clamp(position.row, 0, (worksheet?.rowCount ?? 1) - 1),
      column: clamp(
        position.column,
        0,
        (worksheet?.columnCount ?? 1) - 1,
      ),
    };
  }

  private scrollCellIntoView(position: XlsxCellPosition): void {
    const worksheet = this.worksheet;
    const grid = this.virtualGrid;
    const viewportEl = this.viewportEl;
    if (!worksheet || !grid || !viewportEl) {
      return;
    }
    const scale = this.getEffectiveScale();
    const frozenRows = worksheet.frozenPane?.rows ?? 0;
    const frozenColumns = worksheet.frozenPane?.columns ?? 0;
    if (position.column >= frozenColumns) {
      const left = grid.columnOffset(position.column) * scale;
      const right =
        (grid.columnOffset(position.column) +
          grid.columnSize(position.column)) *
        scale;
      if (left < viewportEl.scrollLeft) {
        viewportEl.scrollLeft = left;
      } else if (right > viewportEl.scrollLeft + viewportEl.clientWidth) {
        viewportEl.scrollLeft = right - viewportEl.clientWidth;
      }
    }
    if (position.row >= frozenRows) {
      const top = grid.rowOffset(position.row) * scale;
      const bottom =
        (grid.rowOffset(position.row) + grid.rowSize(position.row)) * scale;
      if (top < viewportEl.scrollTop) {
        viewportEl.scrollTop = top;
      } else if (bottom > viewportEl.scrollTop + viewportEl.clientHeight) {
        viewportEl.scrollTop = bottom - viewportEl.clientHeight;
      }
    }
  }

  private updateFormulaBar(): void {
    const worksheet = this.worksheet;
    const selection = this.selection;
    const singleCell =
      selection.startRow === selection.endRow &&
      selection.startColumn === selection.endColumn;
    const rangeLabel = singleCell
      ? toCellReference(selection.startRow, selection.startColumn)
      : `${toCellReference(
          selection.startRow,
          selection.startColumn,
        )}:${toCellReference(selection.endRow, selection.endColumn)}`;
    this.formulaNameEl?.setText(rangeLabel);
    this.formulaNameEl?.setAttribute(
      "title",
      this.text.labels.selectedRange(rangeLabel),
    );
    const cell = singleCell
      ? worksheet?.getCell(selection.startRow, selection.startColumn)
      : undefined;
    this.formulaValueEl?.setText(
      cell?.formula ? `=${cell.formula.text}` : cell?.displayValue ?? "",
    );
    this.formulaValueEl?.setAttribute(
      "data-label",
      this.text.labels.formula,
    );
    this.formulaCachedEl?.setText(
      cell?.formula
        ? `${this.text.labels.cachedValue}: ${cell.displayValue}`
        : "",
    );
    this.formulaSafetyEl?.toggleAttribute("hidden", !cell?.formula);
    this.linkButtonEl?.toggleAttribute("hidden", !cell?.hyperlink);
    const focusId = this.getCellElementId(
      this.selectionFocus.row,
      this.selectionFocus.column,
    );
    this.viewportEl?.setAttribute("aria-activedescendant", focusId);
  }

  private async copySelection(mode: XlsxCopyMode): Promise<void> {
    const worksheet = this.worksheet;
    if (!worksheet) {
      return;
    }
    try {
      const value = xlsxSelectionToTsv(worksheet, this.selection, mode);
      await navigator.clipboard.writeText(value);
      new Notice(
        mode === "formula"
          ? this.text.notices.copiedFormulas
          : this.text.notices.copiedValues,
      );
    } catch (error) {
      if (error instanceof XlsxSelectionTooLargeError) {
        new Notice(
          this.text.notices.selectionTooLarge(
            error.cellCount,
            MAX_XLSX_COPY_CELLS,
          ),
        );
        return;
      }
      new Notice(this.text.notices.copyFailed(getErrorMessage(error)));
    }
  }

  private scheduleSearch(): void {
    if (this.searchTimer !== null) {
      this.contentEl.win.clearTimeout(this.searchTimer);
    }
    this.searchTimer = this.contentEl.win.setTimeout(() => {
      this.searchTimer = null;
      void this.runSearch();
    }, SEARCH_DEBOUNCE_MS);
  }

  private async runSearch(): Promise<void> {
    const worksheet = this.worksheet;
    const query = this.searchQuery.trim().toLocaleLowerCase();
    const token = this.searchLifecycle.begin();
    this.searchResults = [];
    this.searchResultKeys.clear();
    this.currentSearchResult = -1;
    this.updateSearchCount();
    this.scheduleGridRender();
    if (!worksheet || !query) {
      this.setReadyStatus();
      return;
    }
    this.setStatus(
      this.text.status.searching(worksheet.name),
      false,
      true,
    );
    try {
      this.searchResults = await searchXlsxCells(
        worksheet.getPopulatedCells(),
        query,
        {
          isCancelled: () =>
            !this.searchLifecycle.isCurrent(token) ||
            this.worksheet !== worksheet,
          yieldControl: () => yieldToWindow(this.contentEl.win),
        },
      );
    } catch (error) {
      if (error instanceof XlsxSearchCancelledError) {
        return;
      }
      throw error;
    }
    this.searchResultKeys = new Set(
      this.searchResults.map((cell) => cellKey(cell.row, cell.column)),
    );
    this.updateSearchCount();
    this.scheduleGridRender();
    this.setReadyStatus();
  }

  private goToSearchResult(direction: -1 | 1): void {
    if (this.searchResults.length === 0) {
      return;
    }
    this.currentSearchResult =
      (
        this.currentSearchResult +
        direction +
        this.searchResults.length
      ) % this.searchResults.length;
    const result = this.searchResults[this.currentSearchResult];
    this.selectionAnchor = result;
    this.selectionFocus = result;
    this.scrollCellIntoView(result);
    this.updateFormulaBar();
    this.updateSearchCount();
    this.scheduleGridRender();
  }

  private updateSearchCount(): void {
    if (!this.searchCountEl) {
      return;
    }
    if (!this.searchQuery.trim()) {
      this.searchCountEl.setText("");
      return;
    }
    const total = this.searchResults.length;
    this.searchCountEl.setText(
      this.currentSearchResult >= 0 && total > 0
        ? `${this.currentSearchResult + 1}/${total}`
        : this.text.status.searchResults(total),
    );
  }

  private async openSelectedHyperlink(): Promise<void> {
    const cell = this.worksheet?.getCell(
      this.selectionFocus.row,
      this.selectionFocus.column,
    );
    if (cell?.hyperlink) {
      await this.openHyperlink(cell.hyperlink);
    }
  }

  private async openHyperlink(hyperlink: XlsxHyperlink): Promise<void> {
    const target = hyperlink.external
      ? hyperlink.target
      : hyperlink.location;
    if (!target) {
      return;
    }
    if (!this.contentEl.win.confirm(this.text.notices.linkConfirmation(target))) {
      return;
    }
    if (!hyperlink.external) {
      await this.navigateWorkbookLocation(target);
      return;
    }
    if (!isSafeXlsxExternalTarget(target)) {
      new Notice(this.text.notices.blockedLink);
      return;
    }
    this.contentEl.win.open(target, "_blank", "noopener,noreferrer");
  }

  private async navigateWorkbookLocation(location: string): Promise<void> {
    const parsed = parseXlsxWorkbookLocation(location);
    const workbook = this.workbook;
    if (!parsed || !workbook) {
      new Notice(this.text.notices.blockedLink);
      return;
    }
    let sheetIndex = this.activeSheetIndex;
    if (parsed.sheetName) {
      const targetIndex = workbook.sheets.findIndex(
        (sheet) => sheet.name === parsed.sheetName,
      );
      if (targetIndex < 0) {
        new Notice(this.text.notices.blockedLink);
        return;
      }
      sheetIndex = targetIndex;
    }
    if (sheetIndex !== this.activeSheetIndex) {
      await this.loadWorksheet(sheetIndex, false);
    }
    const position = this.clampCellPosition(
      parseCellReference(parsed.reference),
    );
    this.selectionAnchor = position;
    this.selectionFocus = position;
    this.scrollCellIntoView(position);
    this.updateFormulaBar();
    this.scheduleGridRender();
  }

  private setZoom(
    value: number,
    anchor?: { left: number; top: number },
  ): void {
    const nextZoom = normalizeZoom(value, {
      min: MIN_ZOOM,
      max: MAX_ZOOM,
      step: ZOOM_STEP,
    });
    const viewportEl = this.viewportEl;
    if (nextZoom === null || !viewportEl) {
      this.updateZoomControl(this.getEffectiveScale());
      return;
    }
    const previousZoom = this.getEffectiveScale();
    const pointer = anchor ?? {
      left: viewportEl.clientWidth / 2,
      top: viewportEl.clientHeight / 2,
    };
    const nextScroll = preserveZoomAnchor(
      {
        left: viewportEl.scrollLeft,
        top: viewportEl.scrollTop,
      },
      pointer,
      previousZoom,
      nextZoom,
    );
    this.fitWidth = false;
    this.zoom = nextZoom;
    this.applyScale();
    viewportEl.scrollLeft = Math.max(0, nextScroll.left);
    viewportEl.scrollTop = Math.max(0, nextScroll.top);
    this.saveReadingState();
  }

  private handleWheelZoom(event: WheelEvent): void {
    const viewportEl = this.viewportEl;
    if (!event.ctrlKey || !viewportEl) {
      return;
    }
    event.preventDefault();
    const rect = viewportEl.getBoundingClientRect();
    const direction = event.deltaY < 0 ? 1 : -1;
    this.setZoom(this.getEffectiveScale() + direction * ZOOM_STEP, {
      left: event.clientX - rect.left,
      top: event.clientY - rect.top,
    });
  }

  private applyScale(): void {
    const scale = this.getEffectiveScale();
    this.fitButtonEl?.toggleClass("is-active", this.fitWidth);
    this.fitButtonEl?.setAttribute(
      "aria-pressed",
      String(this.fitWidth),
    );
    this.updateZoomControl(scale);
    this.scheduleGridRender();
  }

  private getEffectiveScale(): number {
    const grid = this.virtualGrid;
    const viewportEl = this.viewportEl;
    if (!this.fitWidth || !grid || !viewportEl || grid.totalWidth <= 0) {
      return this.zoom;
    }
    return clamp(
      viewportEl.clientWidth / grid.totalWidth,
      MIN_ZOOM,
      MAX_ZOOM,
    );
  }

  private updateZoomControl(scale: number): void {
    if (this.zoomInputEl) {
      this.zoomInputEl.value = String(Math.round(scale * 100));
    }
  }

  private restoreReadingState(file: TFile): void {
    this.activeFilePath = file.path;
    const state = this.plugin.getReadingState(file);
    const defaultZoom = this.plugin.settings.common.defaultZoomPercent / 100;
    this.zoom =
      normalizeZoom(state?.zoom ?? defaultZoom, {
        min: MIN_ZOOM,
        max: MAX_ZOOM,
        step: ZOOM_STEP,
      }) ?? defaultZoom;
    this.fitWidth =
      state?.fitWidth ?? this.plugin.settings.xlsx.defaultFitWidth;
    this.activeSheetIndex = Math.max(0, (state?.page ?? 1) - 1);
    this.pendingScrollPosition = {
      left: state?.scrollLeft ?? 0,
      top: state?.scrollTop ?? 0,
    };
  }

  private restoreScrollPosition(): void {
    const viewportEl = this.viewportEl;
    if (!viewportEl || !this.pendingScrollPosition) {
      return;
    }
    this.renderGrid();
    viewportEl.scrollLeft = this.pendingScrollPosition.left;
    viewportEl.scrollTop = this.pendingScrollPosition.top;
    this.pendingScrollPosition = null;
    this.scheduleGridRender();
  }

  private scheduleStateSave(): void {
    if (this.stateSaveTimer !== null) {
      this.contentEl.win.clearTimeout(this.stateSaveTimer);
    }
    this.stateSaveTimer = this.contentEl.win.setTimeout(() => {
      this.stateSaveTimer = null;
      this.saveReadingState();
    }, 250);
  }

  private saveReadingState(): void {
    if (!this.activeFilePath) {
      return;
    }
    const state: ReaderViewState = {
      zoom: this.zoom,
      fitWidth: this.fitWidth,
      outlineVisible: false,
      scrollLeft: this.viewportEl?.scrollLeft ?? 0,
      scrollTop: this.viewportEl?.scrollTop ?? 0,
      collapsedOutlineIds: [],
      page: this.activeSheetIndex + 1,
    };
    this.plugin.updateReadingState(this.file ?? this.activeFilePath, state);
  }

  private setReadyStatus(): void {
    const file = this.file;
    const worksheet = this.worksheet;
    if (!file || !worksheet) {
      return;
    }
    this.setStatus(
      this.text.status.ready(
        file.name,
        worksheet.name,
        worksheet.rowCount,
        worksheet.columnCount,
      ),
    );
  }

  private showLegacyWorkbook(): void {
    const file = this.file;
    this.clearWorksheet();
    this.clearGridMessage();
    if (!this.gridShellEl || !file) {
      return;
    }
    this.gridMessageEl = this.gridShellEl.createDiv({
      cls: "office-reader-message xlsx-reader-message",
    });
    this.gridMessageEl.createDiv({
      cls: "office-reader-message-title",
      text: this.text.legacy.title,
    });
    this.gridMessageEl.createDiv({
      cls: "office-reader-message-body",
      text: this.text.legacy.body,
    });
    this.gridMessageEl.createDiv({
      cls: "office-reader-message-body",
      text: this.text.legacy.convert,
    });
    const actionsEl = this.gridMessageEl.createDiv({
      cls: "office-reader-message-actions",
    });
    this.createIconButton(
      actionsEl,
      "external-link",
      this.text.toolbar.openExternally,
      () => {
        void this.openExternal();
      },
    );
    this.setStatus(this.text.status.legacy(file.name));
  }

  private showError(error: unknown): void {
    this.clearWorksheet();
    this.clearGridMessage();
    const info = classifyXlsxError(error, this.text);
    if (!this.gridShellEl) {
      return;
    }
    this.gridMessageEl = this.shell.renderError(this.gridShellEl, {
      title: info.title,
      body: info.body,
      tips: info.tips,
      detailsLabel: this.text.errors.details,
      details: info.details,
    });
    this.gridMessageEl.addClass("xlsx-reader-message");
    const actionsEl = this.gridMessageEl.createDiv({
      cls: "office-reader-message-actions",
    });
    this.createIconButton(
      actionsEl,
      "refresh-cw",
      this.text.toolbar.reload,
      () => {
        void this.reload();
      },
    );
    this.createIconButton(
      actionsEl,
      "external-link",
      this.text.toolbar.openExternally,
      () => {
        void this.openExternal();
      },
    );
    this.setStatus(info.title, true);
  }

  private clearGridMessage(): void {
    this.gridMessageEl?.remove();
    this.gridMessageEl = null;
  }

  private updateGridAria(): void {
    this.viewportEl?.setAttribute(
      "aria-rowcount",
      String(this.worksheet?.rowCount ?? 0),
    );
    this.viewportEl?.setAttribute(
      "aria-colcount",
      String(this.worksheet?.columnCount ?? 0),
    );
  }

  private normalizeSheetIndex(
    requestedIndex: number,
    workbook: XlsxPackage,
  ): number {
    if (
      workbook.sheets[requestedIndex]?.state === "visible"
    ) {
      return requestedIndex;
    }
    const firstVisible = workbook.sheets.findIndex(
      (sheet) => sheet.state === "visible",
    );
    return firstVisible >= 0 ? firstVisible : 0;
  }

  private createIconButton(
    parentEl: HTMLElement,
    icon: string,
    label: string,
    onClick: () => void,
  ): HTMLButtonElement {
    const buttonEl = parentEl.createEl("button", {
      cls: "office-reader-toolbar-button",
      attr: {
        type: "button",
        "aria-label": label,
        title: label,
      },
    });
    setIcon(buttonEl, icon);
    buttonEl.addEventListener("click", onClick);
    return buttonEl;
  }

  private createToolbarButton(
    parentEl: HTMLElement,
    capability: ReaderCapability,
    icon: string,
    label: string,
    onClick: () => void,
  ): HTMLButtonElement | null {
    return this.shell.createToolbarButton(
      parentEl,
      capability,
      icon,
      label,
      onClick,
    );
  }

  private getCellElementId(row: number, column: number): string {
    return `${this.cellIdPrefix}-cell-${row}-${column}`;
  }

  private setStatus(
    message: string,
    isError = false,
    isLoading = false,
  ): void {
    this.readerStatus.set(
      message,
      isError ? "error" : isLoading ? "loading" : "idle",
    );
  }

  private renderStatus(status: ReaderStatus): void {
    this.shell.renderStatus(this.statusEl, status);
  }

  private cancelScheduledWork(): void {
    if (this.searchTimer !== null) {
      this.contentEl.win.clearTimeout(this.searchTimer);
      this.searchTimer = null;
    }
    if (this.stateSaveTimer !== null) {
      this.contentEl.win.clearTimeout(this.stateSaveTimer);
      this.stateSaveTimer = null;
    }
    if (this.renderFrameId !== null) {
      this.contentEl.win.cancelAnimationFrame(this.renderFrameId);
      this.renderFrameId = null;
    }
  }
}

function range(start: number, end: number): number[] {
  return Array.from(
    { length: Math.max(0, end - start) },
    (_, index) => start + index,
  );
}

function uniqueSorted(values: readonly number[]): number[] {
  return Array.from(new Set(values)).sort((left, right) => left - right);
}

function cartesian(
  rows: readonly number[],
  columns: readonly number[],
): CellCoordinate[] {
  const coordinates: CellCoordinate[] = [];
  for (const row of rows) {
    for (const column of columns) {
      coordinates.push({ row, column });
    }
  }
  return coordinates;
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function toCellReference(row: number, column: number): string {
  return `${columnIndexToName(column)}${row + 1}`;
}

function mergeContains(
  merge: XlsxMergeRange,
  row: number,
  column: number,
): boolean {
  return (
    row >= merge.startRow &&
    row <= merge.endRow &&
    column >= merge.startColumn &&
    column <= merge.endColumn
  );
}

function mergeIntersectsAxes(
  merge: XlsxMergeRange,
  rows: readonly number[],
  columns: readonly number[],
): boolean {
  return (
    rows.some((row) => row >= merge.startRow && row <= merge.endRow) &&
    columns.some(
      (column) =>
        column >= merge.startColumn && column <= merge.endColumn,
    )
  );
}

function getCellTitle(cell: XlsxCell): string {
  const details = [cell.ref, cell.displayValue];
  if (cell.formula) {
    details.push(`=${cell.formula.text}`);
  }
  if (cell.hyperlink?.tooltip) {
    details.push(cell.hyperlink.tooltip);
  }
  return details.filter(Boolean).join("\n");
}

function yieldToWindow(targetWindow: Window): Promise<void> {
  return new Promise((resolve) => {
    targetWindow.setTimeout(resolve, 0);
  });
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
