import type { CellData, RangeAddress, SheetToJsonConfig, CellValue, DateHandling, TableConfig } from './types';
import type { Workbook } from './workbook';
import { Cell, parseCellRef } from './cell';
import { Range } from './range';
import { Table } from './table';
import { parseRange, toAddress, parseAddress, letterToCol } from './utils/address';
import {
  parseXml,
  findElement,
  getChildren,
  getAttr,
  XmlNode,
  stringifyXml,
  createElement,
  createText,
} from './utils/xml';

/**
 * Represents a worksheet in a workbook
 */
export class Worksheet {
  private _name: string;
  private _workbook: Workbook;
  private _cells: Map<string, Cell> = new Map();
  private _xmlNodes: XmlNode[] | null = null;
  private _dirty = false;
  private _mergedCells: Set<string> = new Set();
  private _sheetData: XmlNode[] = [];
  private _columnWidths: Map<number, number> = new Map();
  private _rowHeights: Map<number, number> = new Map();
  private _frozenPane: { row: number; col: number } | null = null;
  private _dataBoundsCache: { minRow: number; maxRow: number; minCol: number; maxCol: number } | null = null;
  private _boundsDirty = true;
  private _tables: Table[] = [];
  private _preserveXml = false;
  private _tableRelIds: string[] | null = null;
  private _sheetViewsDirty = false;
  private _colsDirty = false;
  private _tablePartsDirty = false;

  constructor(workbook: Workbook, name: string) {
    this._workbook = workbook;
    this._name = name;
  }

  /**
   * Get the workbook this sheet belongs to
   */
  get workbook(): Workbook {
    return this._workbook;
  }

  /**
   * Get the sheet name
   */
  get name(): string {
    return this._name;
  }

  /**
   * Set the sheet name
   */
  set name(value: string) {
    this._name = value;
    this._dirty = true;
  }

  /**
   * Parse worksheet XML content
   */
  parse(xml: string): void {
    this._xmlNodes = parseXml(xml);
    this._preserveXml = true;
    const worksheet = findElement(this._xmlNodes, 'worksheet');
    if (!worksheet) return;

    const worksheetChildren = getChildren(worksheet, 'worksheet');

    // Parse sheet views (freeze panes)
    const sheetViews = findElement(worksheetChildren, 'sheetViews');
    if (sheetViews) {
      const viewChildren = getChildren(sheetViews, 'sheetViews');
      const sheetView = findElement(viewChildren, 'sheetView');
      if (sheetView) {
        const sheetViewChildren = getChildren(sheetView, 'sheetView');
        const pane = findElement(sheetViewChildren, 'pane');
        if (pane && getAttr(pane, 'state') === 'frozen') {
          const xSplit = parseInt(getAttr(pane, 'xSplit') || '0', 10);
          const ySplit = parseInt(getAttr(pane, 'ySplit') || '0', 10);
          if (xSplit > 0 || ySplit > 0) {
            this._frozenPane = { row: ySplit, col: xSplit };
          }
        }
      }
    }

    // Parse sheet data (cells)
    const sheetData = findElement(worksheetChildren, 'sheetData');
    if (sheetData) {
      this._sheetData = getChildren(sheetData, 'sheetData');
      this._parseSheetData(this._sheetData);
    }

    // Parse column widths
    const cols = findElement(worksheetChildren, 'cols');
    if (cols) {
      const colChildren = getChildren(cols, 'cols');
      for (const col of colChildren) {
        if (!('col' in col)) continue;
        const min = parseInt(getAttr(col, 'min') || '0', 10);
        const max = parseInt(getAttr(col, 'max') || '0', 10);
        const width = parseFloat(getAttr(col, 'width') || '0');
        if (!Number.isFinite(width) || width <= 0) continue;
        if (min > 0 && max > 0) {
          for (let idx = min; idx <= max; idx++) {
            this._columnWidths.set(idx - 1, width);
          }
        }
      }
    }

    // Parse merged cells
    const mergeCells = findElement(worksheetChildren, 'mergeCells');
    if (mergeCells) {
      const mergeChildren = getChildren(mergeCells, 'mergeCells');
      for (const mergeCell of mergeChildren) {
        if ('mergeCell' in mergeCell) {
          const ref = getAttr(mergeCell, 'ref');
          if (ref) {
            this._mergedCells.add(ref);
          }
        }
      }
    }
  }

  /**
   * Parse the sheetData element to extract cells
   */
  private _parseSheetData(rows: XmlNode[]): void {
    for (const rowNode of rows) {
      if (!('row' in rowNode)) continue;

      const rowIndex = parseInt(getAttr(rowNode, 'r') || '0', 10) - 1;
      const rowHeight = parseFloat(getAttr(rowNode, 'ht') || '0');
      if (rowIndex >= 0 && Number.isFinite(rowHeight) && rowHeight > 0) {
        this._rowHeights.set(rowIndex, rowHeight);
      }

      const rowChildren = getChildren(rowNode, 'row');
      for (const cellNode of rowChildren) {
        if (!('c' in cellNode)) continue;

        const ref = getAttr(cellNode, 'r');
        if (!ref) continue;

        const { row, col } = parseAddress(ref);
        const cellData = this._parseCellNode(cellNode);
        const cell = new Cell(this, row, col, cellData);
        this._cells.set(ref, cell);
      }
    }

    this._boundsDirty = true;
  }

  /**
   * Parse a cell XML node to CellData
   */
  private _parseCellNode(node: XmlNode): CellData {
    const data: CellData = {};

    // Type attribute
    const t = getAttr(node, 't');
    if (t) {
      data.t = t as CellData['t'];
    }

    // Style attribute
    const s = getAttr(node, 's');
    if (s) {
      data.s = parseInt(s, 10);
    }

    const children = getChildren(node, 'c');

    // Value element
    const vNode = findElement(children, 'v');
    if (vNode) {
      const vChildren = getChildren(vNode, 'v');
      for (const child of vChildren) {
        if ('#text' in child) {
          const text = child['#text'] as string;
          // Parse based on type
          if (data.t === 's') {
            data.v = parseInt(text, 10); // Shared string index
          } else if (data.t === 'b') {
            data.v = text === '1' ? 1 : 0;
          } else if (data.t === 'e' || data.t === 'str') {
            data.v = text;
          } else {
            // Number or default
            data.v = parseFloat(text);
          }
          break;
        }
      }
    }

    // Formula element
    const fNode = findElement(children, 'f');
    if (fNode) {
      const fChildren = getChildren(fNode, 'f');
      for (const child of fChildren) {
        if ('#text' in child) {
          data.f = child['#text'] as string;
          break;
        }
      }

      // Check for shared formula
      const si = getAttr(fNode, 'si');
      if (si) {
        data.si = parseInt(si, 10);
      }

      // Check for array formula range
      const ref = getAttr(fNode, 'ref');
      if (ref) {
        data.F = ref;
      }
    }

    // Inline string (is element)
    const isNode = findElement(children, 'is');
    if (isNode) {
      data.t = 'str';
      const isChildren = getChildren(isNode, 'is');
      const tNode = findElement(isChildren, 't');
      if (tNode) {
        const tChildren = getChildren(tNode, 't');
        for (const child of tChildren) {
          if ('#text' in child) {
            data.v = child['#text'] as string;
            break;
          }
        }
      }
    }

    return data;
  }

  /**
   * Get a cell by address or row/col
   */
  cell(rowOrAddress: number | string, col?: number): Cell {
    const { row, col: c } = parseCellRef(rowOrAddress, col);
    const address = toAddress(row, c);

    let cell = this._cells.get(address);
    if (!cell) {
      cell = new Cell(this, row, c);
      this._cells.set(address, cell);
      this._boundsDirty = true;
    }

    return cell;
  }

  /**
   * Get an existing cell without creating it.
   */
  getCellIfExists(rowOrAddress: number | string, col?: number): Cell | undefined {
    const { row, col: c } = parseCellRef(rowOrAddress, col);
    const address = toAddress(row, c);
    return this._cells.get(address);
  }

  /**
   * Get a range of cells
   */
  range(rangeStr: string): Range;
  range(startRow: number, startCol: number, endRow: number, endCol: number): Range;
  range(startRowOrRange: number | string, startCol?: number, endRow?: number, endCol?: number): Range {
    let rangeAddr: RangeAddress;

    if (typeof startRowOrRange === 'string') {
      rangeAddr = parseRange(startRowOrRange);
    } else {
      if (startCol === undefined || endRow === undefined || endCol === undefined) {
        throw new Error('All range parameters must be provided');
      }
      rangeAddr = {
        start: { row: startRowOrRange, col: startCol },
        end: { row: endRow, col: endCol },
      };
    }

    return new Range(this, rangeAddr);
  }

  /**
   * Merge cells in the given range
   */
  mergeCells(rangeOrStart: string, end?: string): void {
    let rangeStr: string;
    if (end) {
      rangeStr = `${rangeOrStart}:${end}`;
    } else {
      rangeStr = rangeOrStart;
    }
    this._mergedCells.add(rangeStr);
    this._dirty = true;
  }

  /**
   * Unmerge cells in the given range
   */
  unmergeCells(rangeStr: string): void {
    this._mergedCells.delete(rangeStr);
    this._dirty = true;
  }

  /**
   * Get all merged cell ranges
   */
  get mergedCells(): string[] {
    return Array.from(this._mergedCells);
  }

  /**
   * Check if the worksheet has been modified
   */
  get dirty(): boolean {
    if (this._dirty) return true;
    for (const cell of this._cells.values()) {
      if (cell.dirty) return true;
    }
    return false;
  }

  /**
   * Get all cells in the worksheet
   */
  get cells(): Map<string, Cell> {
    return this._cells;
  }

  /**
   * Set a column width (0-based index or column letter)
   */
  setColumnWidth(col: number | string, width: number): void {
    if (!Number.isFinite(width) || width <= 0) {
      throw new Error('Column width must be a positive number');
    }

    const colIndex = typeof col === 'number' ? col : letterToCol(col);
    if (colIndex < 0) {
      throw new Error(`Invalid column: ${col}`);
    }

    this._columnWidths.set(colIndex, width);
    this._colsDirty = true;
    this._dirty = true;
  }

  /**
   * Get a column width if set
   */
  getColumnWidth(col: number | string): number | undefined {
    const colIndex = typeof col === 'number' ? col : letterToCol(col);
    return this._columnWidths.get(colIndex);
  }

  /**
   * Set a row height (0-based index)
   */
  setRowHeight(row: number, height: number): void {
    if (!Number.isFinite(height) || height <= 0) {
      throw new Error('Row height must be a positive number');
    }
    if (row < 0) {
      throw new Error('Row index must be >= 0');
    }

    this._rowHeights.set(row, height);
    this._colsDirty = true;
    this._dirty = true;
  }

  /**
   * Get a row height if set
   */
  getRowHeight(row: number): number | undefined {
    return this._rowHeights.get(row);
  }

  /**
   * Freeze panes at a given row/column split (counts from top-left)
   */
  freezePane(rowSplit: number, colSplit: number): void {
    if (rowSplit < 0 || colSplit < 0) {
      throw new Error('Freeze pane splits must be >= 0');
    }
    if (rowSplit === 0 && colSplit === 0) {
      this._frozenPane = null;
    } else {
      this._frozenPane = { row: rowSplit, col: colSplit };
    }
    this._sheetViewsDirty = true;
    this._dirty = true;
  }

  /**
   * Get current frozen pane configuration
   */
  getFrozenPane(): { row: number; col: number } | null {
    return this._frozenPane ? { ...this._frozenPane } : null;
  }

  /**
   * Get all tables in the worksheet
   */
  get tables(): Table[] {
    return [...this._tables];
  }

  /**
   * Get column width entries
   * @internal
   */
  getColumnWidths(): Map<number, number> {
    return new Map(this._columnWidths);
  }

  /**
   * Get row height entries
   * @internal
   */
  getRowHeights(): Map<number, number> {
    return new Map(this._rowHeights);
  }

  /**
   * Set table relationship IDs for tableParts generation.
   * @internal
   */
  setTableRelIds(ids: string[] | null): void {
    this._tableRelIds = ids ? [...ids] : null;
    this._tablePartsDirty = true;
  }

  /**
   * Create an Excel Table (ListObject) from a data range.
   *
   * Tables provide structured data features like auto-filter, banded styling,
   * and total row with aggregation functions.
   *
   * @param config - Table configuration
   * @returns Table instance for method chaining
   *
   * @example
   * ```typescript
   * // Create a table with default styling
   * const table = sheet.createTable({
   *   name: 'SalesData',
   *   range: 'A1:D10',
   * });
   *
   * // Create a table with total row
   * const table = sheet.createTable({
   *   name: 'SalesData',
   *   range: 'A1:D10',
   *   totalRow: true,
   *   style: { name: 'TableStyleMedium2' }
   * });
   *
   * table.setTotalFunction('Sales', 'sum');
   * ```
   */
  createTable(config: TableConfig): Table {
    // Validate table name is unique within the workbook
    for (const sheet of this._workbook.sheetNames) {
      const ws = this._workbook.sheet(sheet);
      for (const table of ws._tables) {
        if (table.name === config.name) {
          throw new Error(`Table name already exists: ${config.name}`);
        }
      }
    }

    // Validate table name format (Excel rules: no spaces at start/end, alphanumeric + underscore)
    if (!config.name || !/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(config.name)) {
      throw new Error(
        `Invalid table name: ${config.name}. Names must start with a letter or underscore and contain only alphanumeric characters, underscores, or periods.`,
      );
    }

    // Create the table with a unique ID from the workbook
    const tableId = this._workbook.getNextTableId();
    const table = new Table(this, config, tableId);

    this._tables.push(table);
    this._tablePartsDirty = true;
    this._dirty = true;

    return table;
  }

  /**
   * Convert sheet data to an array of JSON objects.
   *
   * @param config - Configuration options
   * @returns Array of objects where keys are field names and values are cell values
   *
   * @example
   * ```typescript
   * // Using first row as headers
   * const data = sheet.toJson();
   *
   * // Using custom field names
   * const data = sheet.toJson({ fields: ['name', 'age', 'city'] });
   *
   * // Starting from a specific row/column
   * const data = sheet.toJson({ startRow: 2, startCol: 1 });
   * ```
   */
  toJson<T = Record<string, CellValue>>(config: SheetToJsonConfig = {}): T[] {
    const {
      fields,
      startRow = 0,
      startCol = 0,
      endRow,
      endCol,
      stopOnEmptyRow = true,
      dateHandling = this._workbook.dateHandling,
      asText = false,
      locale,
    } = config;

    // Get the bounds of data in the sheet
    const bounds = this._getDataBounds();
    if (!bounds) {
      return [];
    }

    const effectiveEndRow = endRow ?? bounds.maxRow;
    const effectiveEndCol = endCol ?? bounds.maxCol;

    // Determine field names
    let fieldNames: string[];
    let dataStartRow: number;

    if (fields) {
      // Use provided field names, data starts at startRow
      fieldNames = fields;
      dataStartRow = startRow;
    } else {
      // Use first row as headers
      fieldNames = [];
      for (let col = startCol; col <= effectiveEndCol; col++) {
        const cell = this._cells.get(toAddress(startRow, col));
        const value = cell?.value;
        fieldNames.push(value != null ? String(value) : `column${col}`);
      }
      dataStartRow = startRow + 1;
    }

    // Read data rows
    const result: T[] = [];

    for (let row = dataStartRow; row <= effectiveEndRow; row++) {
      const obj: Record<string, CellValue | string> = {};
      let hasData = false;

      for (let colOffset = 0; colOffset < fieldNames.length; colOffset++) {
        const col = startCol + colOffset;
        const cell = this._cells.get(toAddress(row, col));

        let value: CellValue | string;

        if (asText) {
          // Return formatted text instead of raw value
          value = cell?.textWithLocale(locale) ?? '';
          if (value !== '') {
            hasData = true;
          }
        } else {
          value = cell?.value ?? null;
          if (value instanceof Date) {
            value = this._serializeDate(value, dateHandling, cell);
          }
          if (value !== null) {
            hasData = true;
          }
        }

        const fieldName = fieldNames[colOffset];
        if (fieldName) {
          obj[fieldName] = value;
        }
      }

      // Stop on empty row if configured
      if (stopOnEmptyRow && !hasData) {
        break;
      }

      result.push(obj as T);
    }

    return result;
  }

  private _serializeDate(value: Date, dateHandling: DateHandling, cell?: Cell | null): CellValue | number | string {
    if (dateHandling === 'excelSerial') {
      return cell?._jsDateToExcel(value) ?? value;
    }

    if (dateHandling === 'isoString') {
      return value.toISOString();
    }

    return value;
  }

  /**
   * Get the bounds of data in the sheet (min/max row and column with data)
   */
  private _getDataBounds(): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null {
    if (!this._boundsDirty && this._dataBoundsCache) {
      return this._dataBoundsCache;
    }

    if (this._cells.size === 0) {
      this._dataBoundsCache = null;
      this._boundsDirty = false;
      return null;
    }

    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;

    for (const cell of this._cells.values()) {
      if (cell.value !== null) {
        minRow = Math.min(minRow, cell.row);
        maxRow = Math.max(maxRow, cell.row);
        minCol = Math.min(minCol, cell.col);
        maxCol = Math.max(maxCol, cell.col);
      }
    }

    if (minRow === Infinity) {
      this._dataBoundsCache = null;
      this._boundsDirty = false;
      return null;
    }

    this._dataBoundsCache = { minRow, maxRow, minCol, maxCol };
    this._boundsDirty = false;
    return this._dataBoundsCache;
  }

  /**
   * Generate XML for this worksheet
   */
  toXml(): string {
    const preserved = this._preserveXml && this._xmlNodes ? this._buildPreservedWorksheet() : null;
    // Build sheetData from cells
    const sheetDataNode = this._buildSheetDataNode();

    // Build worksheet structure
    const worksheetChildren: XmlNode[] = [];

    // Sheet views (freeze panes)
    if (this._frozenPane) {
      const paneAttrs: Record<string, string> = { state: 'frozen' };
      const topLeftCell = toAddress(this._frozenPane.row, this._frozenPane.col);
      paneAttrs.topLeftCell = topLeftCell;
      if (this._frozenPane.col > 0) {
        paneAttrs.xSplit = String(this._frozenPane.col);
      }
      if (this._frozenPane.row > 0) {
        paneAttrs.ySplit = String(this._frozenPane.row);
      }

      let activePane = 'bottomRight';
      if (this._frozenPane.row > 0 && this._frozenPane.col === 0) {
        activePane = 'bottomLeft';
      } else if (this._frozenPane.row === 0 && this._frozenPane.col > 0) {
        activePane = 'topRight';
      }

      paneAttrs.activePane = activePane;
      const paneNode = createElement('pane', paneAttrs, []);
      const selectionNode = createElement(
        'selection',
        { pane: activePane, activeCell: topLeftCell, sqref: topLeftCell },
        [],
      );

      const sheetViewNode = createElement('sheetView', { workbookViewId: '0' }, [paneNode, selectionNode]);
      worksheetChildren.push(createElement('sheetViews', {}, [sheetViewNode]));
    }

    // Column widths
    if (this._columnWidths.size > 0) {
      const colNodes: XmlNode[] = [];
      const entries = Array.from(this._columnWidths.entries()).sort((a, b) => a[0] - b[0]);
      for (const [colIndex, width] of entries) {
        colNodes.push(
          createElement(
            'col',
            {
              min: String(colIndex + 1),
              max: String(colIndex + 1),
              width: String(width),
              customWidth: '1',
            },
            [],
          ),
        );
      }
      worksheetChildren.push(createElement('cols', {}, colNodes));
    }

    worksheetChildren.push(sheetDataNode);

    // Add merged cells if any
    if (this._mergedCells.size > 0) {
      const mergeCellNodes: XmlNode[] = [];
      for (const ref of this._mergedCells) {
        mergeCellNodes.push(createElement('mergeCell', { ref }, []));
      }
      const mergeCellsNode = createElement('mergeCells', { count: String(this._mergedCells.size) }, mergeCellNodes);
      worksheetChildren.push(mergeCellsNode);
    }

    // Add table parts if any tables exist
    const tablePartsNode = this._buildTablePartsNode();
    if (tablePartsNode) {
      worksheetChildren.push(tablePartsNode);
    }

    const worksheetNode = createElement(
      'worksheet',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      },
      worksheetChildren,
    );

    if (preserved) {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([preserved])}`;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([worksheetNode])}`;
  }

  private _buildSheetDataNode(): XmlNode {
    const rowMap = new Map<number, Cell[]>();
    for (const cell of this._cells.values()) {
      const row = cell.row;
      if (!rowMap.has(row)) {
        rowMap.set(row, []);
      }
      rowMap.get(row)!.push(cell);
    }

    for (const rowIdx of this._rowHeights.keys()) {
      if (!rowMap.has(rowIdx)) {
        rowMap.set(rowIdx, []);
      }
    }

    const sortedRows = Array.from(rowMap.entries()).sort((a, b) => a[0] - b[0]);
    const rowNodes: XmlNode[] = [];
    for (const [rowIdx, cells] of sortedRows) {
      cells.sort((a, b) => a.col - b.col);

      const cellNodes: XmlNode[] = [];
      for (const cell of cells) {
        const cellNode = this._buildCellNode(cell);
        cellNodes.push(cellNode);
      }

      const rowAttrs: Record<string, string> = { r: String(rowIdx + 1) };
      const rowHeight = this._rowHeights.get(rowIdx);
      if (rowHeight !== undefined) {
        rowAttrs.ht = String(rowHeight);
        rowAttrs.customHeight = '1';
      }
      const rowNode = createElement('row', rowAttrs, cellNodes);
      rowNodes.push(rowNode);
    }

    return createElement('sheetData', {}, rowNodes);
  }

  private _buildSheetViewsNode(): XmlNode | null {
    if (!this._frozenPane) return null;
    const paneAttrs: Record<string, string> = { state: 'frozen' };
    const topLeftCell = toAddress(this._frozenPane.row, this._frozenPane.col);
    paneAttrs.topLeftCell = topLeftCell;
    if (this._frozenPane.col > 0) {
      paneAttrs.xSplit = String(this._frozenPane.col);
    }
    if (this._frozenPane.row > 0) {
      paneAttrs.ySplit = String(this._frozenPane.row);
    }

    let activePane = 'bottomRight';
    if (this._frozenPane.row > 0 && this._frozenPane.col === 0) {
      activePane = 'bottomLeft';
    } else if (this._frozenPane.row === 0 && this._frozenPane.col > 0) {
      activePane = 'topRight';
    }

    paneAttrs.activePane = activePane;
    const paneNode = createElement('pane', paneAttrs, []);
    const selectionNode = createElement(
      'selection',
      { pane: activePane, activeCell: topLeftCell, sqref: topLeftCell },
      [],
    );

    const sheetViewNode = createElement('sheetView', { workbookViewId: '0' }, [paneNode, selectionNode]);
    return createElement('sheetViews', {}, [sheetViewNode]);
  }

  private _buildColsNode(): XmlNode | null {
    if (this._columnWidths.size === 0) return null;
    const colNodes: XmlNode[] = [];
    const entries = Array.from(this._columnWidths.entries()).sort((a, b) => a[0] - b[0]);
    for (const [colIndex, width] of entries) {
      colNodes.push(
        createElement(
          'col',
          {
            min: String(colIndex + 1),
            max: String(colIndex + 1),
            width: String(width),
            customWidth: '1',
          },
          [],
        ),
      );
    }
    return createElement('cols', {}, colNodes);
  }

  private _buildMergeCellsNode(): XmlNode | null {
    if (this._mergedCells.size === 0) return null;
    const mergeCellNodes: XmlNode[] = [];
    for (const ref of this._mergedCells) {
      mergeCellNodes.push(createElement('mergeCell', { ref }, []));
    }
    return createElement('mergeCells', { count: String(this._mergedCells.size) }, mergeCellNodes);
  }

  private _buildTablePartsNode(): XmlNode | null {
    if (this._tables.length === 0) return null;
    const tablePartNodes: XmlNode[] = [];
    for (let i = 0; i < this._tables.length; i++) {
      const relId =
        this._tableRelIds && this._tableRelIds.length === this._tables.length ? this._tableRelIds[i] : `rId${i + 1}`;
      tablePartNodes.push(createElement('tablePart', { 'r:id': relId }, []));
    }
    return createElement('tableParts', { count: String(this._tables.length) }, tablePartNodes);
  }

  private _buildPreservedWorksheet(): XmlNode | null {
    if (!this._xmlNodes) return null;
    const worksheet = findElement(this._xmlNodes, 'worksheet');
    if (!worksheet) return null;

    const children = getChildren(worksheet, 'worksheet');

    const upsertChild = (tag: string, node: XmlNode | null) => {
      const existingIndex = children.findIndex((child) => tag in child);
      if (node) {
        if (existingIndex >= 0) {
          children[existingIndex] = node;
        } else {
          children.push(node);
        }
      } else if (existingIndex >= 0) {
        children.splice(existingIndex, 1);
      }
    };

    if (this._sheetViewsDirty) {
      const sheetViewsNode = this._buildSheetViewsNode();
      upsertChild('sheetViews', sheetViewsNode);
    }

    if (this._colsDirty) {
      const colsNode = this._buildColsNode();
      upsertChild('cols', colsNode);
    }

    const sheetDataNode = this._buildSheetDataNode();
    upsertChild('sheetData', sheetDataNode);

    const mergeCellsNode = this._buildMergeCellsNode();
    upsertChild('mergeCells', mergeCellsNode);

    if (this._tablePartsDirty) {
      const tablePartsNode = this._buildTablePartsNode();
      upsertChild('tableParts', tablePartsNode);
    }

    return worksheet;
  }

  /**
   * Build a cell XML node from a Cell object
   */
  private _buildCellNode(cell: Cell): XmlNode {
    const data = cell.data;
    const attrs: Record<string, string> = { r: cell.address };

    if (data.t && data.t !== 'n') {
      attrs.t = data.t;
    }
    if (data.s !== undefined) {
      attrs.s = String(data.s);
    }

    const children: XmlNode[] = [];

    // Formula
    if (data.f) {
      const fAttrs: Record<string, string> = {};
      if (data.F) fAttrs.ref = data.F;
      if (data.si !== undefined) fAttrs.si = String(data.si);
      children.push(createElement('f', fAttrs, [createText(data.f)]));
    }

    // Value
    if (data.v !== undefined) {
      children.push(createElement('v', {}, [createText(String(data.v))]));
    }

    return createElement('c', attrs, children);
  }
}
