import type { CellValue, CellStyle, RangeAddress } from './types';
import type { Worksheet } from './worksheet';
import { toAddress, normalizeRange } from './utils/address';

/**
 * Represents a range of cells in a worksheet
 */
export class Range {
  private _worksheet: Worksheet;
  private _range: RangeAddress;

  constructor(worksheet: Worksheet, range: RangeAddress) {
    this._worksheet = worksheet;
    this._range = normalizeRange(range);
  }

  /**
   * Get the range address as a string
   */
  get address(): string {
    const start = toAddress(this._range.start.row, this._range.start.col);
    const end = toAddress(this._range.end.row, this._range.end.col);
    if (start === end) return start;
    return `${start}:${end}`;
  }

  /**
   * Get the number of rows in the range
   */
  get rowCount(): number {
    return this._range.end.row - this._range.start.row + 1;
  }

  /**
   * Get the number of columns in the range
   */
  get colCount(): number {
    return this._range.end.col - this._range.start.col + 1;
  }

  /**
   * Get all values in the range as a 2D array
   */
  get values(): CellValue[][] {
    const result: CellValue[][] = [];
    for (let r = this._range.start.row; r <= this._range.end.row; r++) {
      const row: CellValue[] = [];
      for (let c = this._range.start.col; c <= this._range.end.col; c++) {
        const cell = this._worksheet.cell(r, c);
        row.push(cell.value);
      }
      result.push(row);
    }
    return result;
  }

  /**
   * Set values in the range from a 2D array
   */
  set values(data: CellValue[][]) {
    for (let r = 0; r < data.length && r < this.rowCount; r++) {
      const row = data[r];
      for (let c = 0; c < row.length && c < this.colCount; c++) {
        const cell = this._worksheet.cell(this._range.start.row + r, this._range.start.col + c);
        cell.value = row[c];
      }
    }
  }

  /**
   * Get all formulas in the range as a 2D array
   */
  get formulas(): (string | undefined)[][] {
    const result: (string | undefined)[][] = [];
    for (let r = this._range.start.row; r <= this._range.end.row; r++) {
      const row: (string | undefined)[] = [];
      for (let c = this._range.start.col; c <= this._range.end.col; c++) {
        const cell = this._worksheet.cell(r, c);
        row.push(cell.formula);
      }
      result.push(row);
    }
    return result;
  }

  /**
   * Set formulas in the range from a 2D array
   */
  set formulas(data: (string | undefined)[][]) {
    for (let r = 0; r < data.length && r < this.rowCount; r++) {
      const row = data[r];
      for (let c = 0; c < row.length && c < this.colCount; c++) {
        const cell = this._worksheet.cell(this._range.start.row + r, this._range.start.col + c);
        cell.formula = row[c];
      }
    }
  }

  /**
   * Get the style of the top-left cell
   */
  get style(): CellStyle {
    return this._worksheet.cell(this._range.start.row, this._range.start.col).style;
  }

  /**
   * Set style for all cells in the range
   */
  set style(style: CellStyle) {
    for (let r = this._range.start.row; r <= this._range.end.row; r++) {
      for (let c = this._range.start.col; c <= this._range.end.col; c++) {
        const cell = this._worksheet.cell(r, c);
        cell.style = style;
      }
    }
  }

  /**
   * Iterate over all cells in the range
   */
  *[Symbol.iterator]() {
    for (let r = this._range.start.row; r <= this._range.end.row; r++) {
      for (let c = this._range.start.col; c <= this._range.end.col; c++) {
        yield this._worksheet.cell(r, c);
      }
    }
  }

  /**
   * Iterate over cells row by row
   */
  *rows() {
    for (let r = this._range.start.row; r <= this._range.end.row; r++) {
      const row = [];
      for (let c = this._range.start.col; c <= this._range.end.col; c++) {
        row.push(this._worksheet.cell(r, c));
      }
      yield row;
    }
  }
}
