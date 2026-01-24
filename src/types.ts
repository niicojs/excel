/**
 * Cell value types - what a cell can contain
 */
export type CellValue = number | string | boolean | Date | null | CellError;

/**
 * Represents an Excel error value
 */
export interface CellError {
  error: ErrorType;
}

export type ErrorType =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA';

/**
 * Discriminator for cell content type
 */
export type CellType = 'number' | 'string' | 'boolean' | 'date' | 'error' | 'empty';

/**
 * Style definition for cells
 */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'single' | 'double';
  strike?: boolean;
  fontSize?: number;
  fontName?: string;
  fontColor?: string;
  fill?: string;
  border?: BorderStyle;
  alignment?: Alignment;
  numberFormat?: string;
}

export interface BorderStyle {
  top?: BorderType;
  bottom?: BorderType;
  left?: BorderType;
  right?: BorderType;
}

export type BorderType = 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';

export interface Alignment {
  horizontal?: 'left' | 'center' | 'right' | 'justify';
  vertical?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  textRotation?: number;
}

/**
 * Cell address with 0-indexed row and column
 */
export interface CellAddress {
  row: number;
  col: number;
}

/**
 * Range address with start and end cells
 */
export interface RangeAddress {
  start: CellAddress;
  end: CellAddress;
}

/**
 * Internal cell data representation
 */
export interface CellData {
  /** Cell type: n=number, s=string (shared), str=inline string, b=boolean, e=error, d=date */
  t?: 'n' | 's' | 'str' | 'b' | 'e' | 'd';
  /** Raw value */
  v?: number | string | boolean;
  /** Formula (without leading =) */
  f?: string;
  /** Style index */
  s?: number;
  /** Formatted text (cached) */
  w?: string;
  /** Number format */
  z?: string;
  /** Array formula range */
  F?: string;
  /** Dynamic array formula flag */
  D?: boolean;
  /** Shared formula index */
  si?: number;
}

/**
 * Sheet definition from workbook.xml
 */
export interface SheetDefinition {
  name: string;
  sheetId: number;
  rId: string;
}

/**
 * Relationship definition
 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}
