import type { TableConfig, TableStyleConfig, TableTotalFunction, RangeAddress } from './types';
import type { Worksheet } from './worksheet';
import { parseRange, toAddress, toRange } from './utils/address';
import { createElement, stringifyXml, XmlNode } from './utils/xml';

/**
 * Maps table total function names to SUBTOTAL function numbers
 * SUBTOTAL uses 101-111 for functions that ignore hidden values
 */
const TOTAL_FUNCTION_NUMBERS: Record<TableTotalFunction, number> = {
  average: 101,
  count: 102,
  countNums: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
  none: 0,
};

/**
 * Maps table total function names to XML attribute values
 */
const TOTAL_FUNCTION_NAMES: Record<TableTotalFunction, string> = {
  average: 'average',
  count: 'count',
  countNums: 'countNums',
  max: 'max',
  min: 'min',
  stdDev: 'stdDev',
  sum: 'sum',
  var: 'var',
  none: 'none',
};

/**
 * Represents an Excel Table (ListObject) with auto-filter, banded styling, and total row.
 */
export class Table {
  private _name: string;
  private _displayName: string;
  private _worksheet: Worksheet;
  private _range: RangeAddress;
  private _totalRow: boolean;
  private _autoFilter: boolean;
  private _style: TableStyleConfig;
  private _columns: TableColumn[] = [];
  private _id: number;
  private _dirty = true;

  constructor(worksheet: Worksheet, config: TableConfig, tableId: number) {
    this._worksheet = worksheet;
    this._name = config.name;
    this._displayName = config.name;
    this._range = parseRange(config.range);
    this._totalRow = config.totalRow === true; // Default false
    this._autoFilter = true; // Tables have auto-filter by default
    this._id = tableId;

    // Expand range to include total row if enabled
    if (this._totalRow) {
      this._range.end.row++;
    }

    // Set default style
    this._style = {
      name: config.style?.name ?? 'TableStyleMedium2',
      showRowStripes: config.style?.showRowStripes !== false, // Default true
      showColumnStripes: config.style?.showColumnStripes === true, // Default false
      showFirstColumn: config.style?.showFirstColumn === true, // Default false
      showLastColumn: config.style?.showLastColumn === true, // Default false
    };

    // Extract column names from worksheet headers
    this._extractColumns();
  }

  /**
   * Get the table name
   */
  get name(): string {
    return this._name;
  }

  /**
   * Get the table display name
   */
  get displayName(): string {
    return this._displayName;
  }

  /**
   * Get the table ID
   */
  get id(): number {
    return this._id;
  }

  /**
   * Get the worksheet this table belongs to
   */
  get worksheet(): Worksheet {
    return this._worksheet;
  }

  /**
   * Get the table range address string
   */
  get range(): string {
    return toRange(this._range);
  }

  /**
   * Get the table range as RangeAddress
   */
  get rangeAddress(): RangeAddress {
    return { ...this._range };
  }

  /**
   * Get column names
   */
  get columns(): string[] {
    return this._columns.map((c) => c.name);
  }

  /**
   * Check if table has a total row
   */
  get hasTotalRow(): boolean {
    return this._totalRow;
  }

  /**
   * Check if table has auto-filter enabled
   */
  get hasAutoFilter(): boolean {
    return this._autoFilter;
  }

  /**
   * Get the current style configuration
   */
  get style(): TableStyleConfig {
    return { ...this._style };
  }

  /**
   * Check if the table has been modified
   */
  get dirty(): boolean {
    return this._dirty;
  }

  /**
   * Set a total function for a column
   * @param columnName - Name of the column (header text)
   * @param fn - Aggregation function to use
   * @returns this for method chaining
   */
  setTotalFunction(columnName: string, fn: TableTotalFunction): this {
    if (!this._totalRow) {
      throw new Error('Cannot set total function: table does not have a total row enabled');
    }

    const column = this._columns.find((c) => c.name === columnName);
    if (!column) {
      throw new Error(`Column not found: ${columnName}`);
    }

    column.totalFunction = fn;
    this._dirty = true;

    // Write the formula to the total row cell
    this._writeTotalRowFormula(column);

    return this;
  }

  /**
   * Enable or disable auto-filter
   * @param enabled - Whether auto-filter should be enabled
   * @returns this for method chaining
   */
  setAutoFilter(enabled: boolean): this {
    this._autoFilter = enabled;
    this._dirty = true;
    return this;
  }

  /**
   * Update table style configuration
   * @param style - Style options to apply
   * @returns this for method chaining
   */
  setStyle(style: Partial<TableStyleConfig>): this {
    if (style.name !== undefined) this._style.name = style.name;
    if (style.showRowStripes !== undefined) this._style.showRowStripes = style.showRowStripes;
    if (style.showColumnStripes !== undefined) this._style.showColumnStripes = style.showColumnStripes;
    if (style.showFirstColumn !== undefined) this._style.showFirstColumn = style.showFirstColumn;
    if (style.showLastColumn !== undefined) this._style.showLastColumn = style.showLastColumn;
    this._dirty = true;
    return this;
  }

  /**
   * Enable or disable the total row
   * @param enabled - Whether total row should be shown
   * @returns this for method chaining
   */
  setTotalRow(enabled: boolean): this {
    if (this._totalRow === enabled) return this;

    this._totalRow = enabled;
    this._dirty = true;

    if (enabled) {
      // Expand range to include total row
      this._range.end.row++;
    } else {
      // Contract range to exclude total row (if it was added)
      // Clear total functions
      for (const col of this._columns) {
        col.totalFunction = undefined;
      }
    }

    return this;
  }

  /**
   * Extract column names from the header row of the worksheet
   */
  private _extractColumns(): void {
    const headerRow = this._range.start.row;
    const startCol = this._range.start.col;
    const endCol = this._range.end.col;

    for (let col = startCol; col <= endCol; col++) {
      const cell = this._worksheet.getCellIfExists(headerRow, col);
      const value = cell?.value;
      const name = value != null ? String(value) : `Column${col - startCol + 1}`;

      this._columns.push({
        id: col - startCol + 1,
        name,
        colIndex: col,
      });
    }
  }

  /**
   * Write the SUBTOTAL formula to a total row cell
   */
  private _writeTotalRowFormula(column: TableColumn): void {
    if (!this._totalRow || !column.totalFunction || column.totalFunction === 'none') {
      return;
    }

    const totalRowIndex = this._range.end.row;
    const cell = this._worksheet.cell(totalRowIndex, column.colIndex);

    // Generate SUBTOTAL formula with structured reference
    const funcNum = TOTAL_FUNCTION_NUMBERS[column.totalFunction];
    // Use structured reference: SUBTOTAL(109,[ColumnName])
    const formula = `SUBTOTAL(${funcNum},[${column.name}])`;
    cell.formula = formula;
  }

  /**
   * Get the auto-filter range (excludes total row if present)
   */
  private _getAutoFilterRange(): string {
    const start = toAddress(this._range.start.row, this._range.start.col);

    // Auto-filter excludes the total row
    let endRow = this._range.end.row;
    if (this._totalRow) {
      endRow--;
    }

    const end = toAddress(endRow, this._range.end.col);
    return `${start}:${end}`;
  }

  /**
   * Generate the table definition XML
   */
  toXml(): string {
    const children: XmlNode[] = [];

    // Auto-filter element
    if (this._autoFilter) {
      const autoFilterRef = this._getAutoFilterRange();
      children.push(createElement('autoFilter', { ref: autoFilterRef }, []));
    }

    // Table columns
    const columnNodes: XmlNode[] = this._columns.map((col) => {
      const attrs: Record<string, string> = {
        id: String(col.id),
        name: col.name,
      };

      // Add total function if specified
      if (this._totalRow && col.totalFunction && col.totalFunction !== 'none') {
        attrs.totalsRowFunction = TOTAL_FUNCTION_NAMES[col.totalFunction];
      }

      return createElement('tableColumn', attrs, []);
    });

    children.push(createElement('tableColumns', { count: String(columnNodes.length) }, columnNodes));

    // Table style info
    const styleAttrs: Record<string, string> = {
      name: this._style.name || 'TableStyleMedium2',
      showFirstColumn: this._style.showFirstColumn ? '1' : '0',
      showLastColumn: this._style.showLastColumn ? '1' : '0',
      showRowStripes: this._style.showRowStripes !== false ? '1' : '0',
      showColumnStripes: this._style.showColumnStripes ? '1' : '0',
    };
    children.push(createElement('tableStyleInfo', styleAttrs, []));

    // Build table attributes
    const tableRef = toRange(this._range);
    const tableAttrs: Record<string, string> = {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      id: String(this._id),
      name: this._name,
      displayName: this._displayName,
      ref: tableRef,
    };

    if (this._totalRow) {
      tableAttrs.totalsRowCount = '1';
    } else {
      tableAttrs.totalsRowShown = '0';
    }

    // Build complete table node
    const tableNode = createElement('table', tableAttrs, children);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([tableNode])}`;
  }
}

/**
 * Internal column representation
 */
interface TableColumn {
  id: number;
  name: string;
  colIndex: number;
  totalFunction?: TableTotalFunction;
}
