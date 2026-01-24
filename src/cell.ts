import type { CellValue, CellType, CellStyle, CellData, ErrorType } from './types';
import type { Worksheet } from './worksheet';
import { parseAddress, toAddress } from './utils/address';

// Excel epoch: December 30, 1899 (accounting for the 1900 leap year bug)
const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30));
const MS_PER_DAY = 24 * 60 * 60 * 1000;

// Excel error types
const ERROR_TYPES: Set<string> = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
]);

/**
 * Represents a single cell in a worksheet
 */
export class Cell {
  private _row: number;
  private _col: number;
  private _data: CellData;
  private _worksheet: Worksheet;
  private _dirty = false;

  constructor(worksheet: Worksheet, row: number, col: number, data?: CellData) {
    this._worksheet = worksheet;
    this._row = row;
    this._col = col;
    this._data = data || {};
  }

  /**
   * Get the cell address (e.g., 'A1')
   */
  get address(): string {
    return toAddress(this._row, this._col);
  }

  /**
   * Get the 0-based row index
   */
  get row(): number {
    return this._row;
  }

  /**
   * Get the 0-based column index
   */
  get col(): number {
    return this._col;
  }

  /**
   * Get the cell type
   */
  get type(): CellType {
    const t = this._data.t;
    if (!t && this._data.v === undefined && !this._data.f) {
      return 'empty';
    }
    switch (t) {
      case 'n':
        return 'number';
      case 's':
      case 'str':
        return 'string';
      case 'b':
        return 'boolean';
      case 'e':
        return 'error';
      case 'd':
        return 'date';
      default:
        // If no type but has value, infer from value
        if (typeof this._data.v === 'number') return 'number';
        if (typeof this._data.v === 'string') return 'string';
        if (typeof this._data.v === 'boolean') return 'boolean';
        return 'empty';
    }
  }

  /**
   * Get the cell value
   */
  get value(): CellValue {
    const t = this._data.t;
    const v = this._data.v;

    if (v === undefined && !this._data.f) {
      return null;
    }

    switch (t) {
      case 'n':
        return typeof v === 'number' ? v : parseFloat(String(v));
      case 's':
        // Shared string reference
        if (typeof v === 'number') {
          return this._worksheet.workbook.sharedStrings.getString(v) ?? '';
        }
        return String(v);
      case 'str':
        // Inline string
        return String(v);
      case 'b':
        return v === 1 || v === '1' || v === true;
      case 'e':
        return { error: String(v) as ErrorType };
      case 'd':
        // ISO 8601 date string
        return new Date(String(v));
      default:
        // No type specified - try to infer
        if (typeof v === 'number') {
          // Check if this might be a date based on number format
          if (this._isDateFormat()) {
            return this._excelDateToJs(v);
          }
          return v;
        }
        if (typeof v === 'string') {
          if (ERROR_TYPES.has(v)) {
            return { error: v as ErrorType };
          }
          return v;
        }
        if (typeof v === 'boolean') return v;
        return null;
    }
  }

  /**
   * Set the cell value
   */
  set value(val: CellValue) {
    this._dirty = true;

    if (val === null || val === undefined) {
      this._data.v = undefined;
      this._data.t = undefined;
      this._data.f = undefined;
      return;
    }

    if (typeof val === 'number') {
      this._data.v = val;
      this._data.t = 'n';
    } else if (typeof val === 'string') {
      // Store as shared string
      const index = this._worksheet.workbook.sharedStrings.addString(val);
      this._data.v = index;
      this._data.t = 's';
    } else if (typeof val === 'boolean') {
      this._data.v = val ? 1 : 0;
      this._data.t = 'b';
    } else if (val instanceof Date) {
      // Store as ISO date string with 'd' type
      this._data.v = val.toISOString();
      this._data.t = 'd';
    } else if ('error' in val) {
      this._data.v = val.error;
      this._data.t = 'e';
    }

    // Clear formula when setting value directly
    this._data.f = undefined;
  }

  /**
   * Write a 2D array of values starting at this cell
   */
  set values(data: CellValue[][]) {
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      for (let c = 0; c < row.length; c++) {
        const cell = this._worksheet.cell(this._row + r, this._col + c);
        cell.value = row[c];
      }
    }
  }

  /**
   * Get the formula (without leading '=')
   */
  get formula(): string | undefined {
    return this._data.f;
  }

  /**
   * Set the formula (without leading '=')
   */
  set formula(f: string | undefined) {
    this._dirty = true;
    if (f === undefined) {
      this._data.f = undefined;
    } else {
      // Remove leading '=' if present
      this._data.f = f.startsWith('=') ? f.slice(1) : f;
    }
  }

  /**
   * Get the formatted text (as displayed in Excel)
   */
  get text(): string {
    if (this._data.w) {
      return this._data.w;
    }
    const val = this.value;
    if (val === null) return '';
    if (typeof val === 'object' && 'error' in val) return val.error;
    if (val instanceof Date) return val.toISOString().split('T')[0];
    return String(val);
  }

  /**
   * Get the style index
   */
  get styleIndex(): number | undefined {
    return this._data.s;
  }

  /**
   * Set the style index
   */
  set styleIndex(index: number | undefined) {
    this._dirty = true;
    this._data.s = index;
  }

  /**
   * Get the cell style
   */
  get style(): CellStyle {
    if (this._data.s === undefined) {
      return {};
    }
    return this._worksheet.workbook.styles.getStyle(this._data.s);
  }

  /**
   * Set the cell style (merges with existing)
   */
  set style(style: CellStyle) {
    this._dirty = true;
    const currentStyle = this.style;
    const merged = { ...currentStyle, ...style };
    this._data.s = this._worksheet.workbook.styles.createStyle(merged);
  }

  /**
   * Check if cell has been modified
   */
  get dirty(): boolean {
    return this._dirty;
  }

  /**
   * Get internal cell data
   */
  get data(): CellData {
    return this._data;
  }

  /**
   * Check if this cell has a date number format
   */
  private _isDateFormat(): boolean {
    // TODO: Check actual number format from styles
    // For now, return false - dates should be explicitly typed
    return false;
  }

  /**
   * Convert Excel serial date to JavaScript Date
   * Used when reading dates stored as numbers with date formats
   */
  _excelDateToJs(serial: number): Date {
    // Excel incorrectly considers 1900 a leap year
    // Dates after Feb 28, 1900 need adjustment
    const adjusted = serial > 60 ? serial - 1 : serial;
    const ms = Math.round((adjusted - 1) * MS_PER_DAY);
    return new Date(EXCEL_EPOCH.getTime() + ms);
  }

  /**
   * Convert JavaScript Date to Excel serial date
   * Used when writing dates as numbers for Excel compatibility
   */
  _jsDateToExcel(date: Date): number {
    const ms = date.getTime() - EXCEL_EPOCH.getTime();
    let serial = ms / MS_PER_DAY + 1;
    // Account for Excel's 1900 leap year bug
    if (serial > 60) {
      serial += 1;
    }
    return serial;
  }
}

/**
 * Parse a cell address or row/col to get row and col indices
 */
export const parseCellRef = (rowOrAddress: number | string, col?: number): { row: number; col: number } => {
  if (typeof rowOrAddress === 'string') {
    return parseAddress(rowOrAddress);
  }
  if (col === undefined) {
    throw new Error('Column must be provided when row is a number');
  }
  return { row: rowOrAddress, col };
};
