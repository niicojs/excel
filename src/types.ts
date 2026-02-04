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

export type ErrorType = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA';

/**
 * Discriminator for cell content type
 */
export type CellType = 'number' | 'string' | 'boolean' | 'date' | 'error' | 'empty';

/**
 * Date handling strategy when serializing cell values.
 */
export type DateHandling = 'jsDate' | 'excelSerial' | 'isoString';

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
  fontColorTheme?: number;
  fontColorTint?: number;
  fontColorIndexed?: number;
  fill?: string;
  fillTheme?: number;
  fillTint?: number;
  fillIndexed?: number;
  fillBgColor?: string;
  fillBgTheme?: number;
  fillBgTint?: number;
  fillBgIndexed?: number;
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

/**
 * Pivot table aggregation functions
 */
export type AggregationType = 'sum' | 'count' | 'average' | 'min' | 'max';

/**
 * Sort order for pivot fields.
 */
export type PivotSortOrder = 'asc' | 'desc';

/**
 * Filter configuration for pivot fields.
 */
export interface PivotFieldFilter {
  include?: string[];
  exclude?: string[];
}

/**
 * Configuration for a value field in a pivot table
 */
export interface PivotValueConfig {
  /** Source field name (column header) */
  field: string;
  /** Aggregation function (default: 'sum') */
  aggregation?: AggregationType;
  /** Display name (e.g., "Sum of Sales") */
  name?: string;
  /** Number format (e.g., '$#,##0.00', '0.00%') */
  numberFormat?: string;
}

/**
 * Configuration for creating a pivot table
 */
export interface PivotTableConfig {
  /** Name of the pivot table */
  name: string;
  /** Source data range with sheet name (e.g., "Sheet1!A1:D100") */
  source: string;
  /** Target cell where pivot table will be placed (e.g., "Sheet2!A3") */
  target: string;
  /** Refresh the pivot table data when the file is opened (default: true) */
  refreshOnLoad?: boolean;
}

/**
 * Internal representation of a pivot cache field
 */
export interface PivotCacheField {
  /** Field name (from header row) */
  name: string;
  /** Field index (0-based) */
  index: number;
  /** Whether this field contains numbers */
  isNumeric: boolean;
  /** Whether this field contains dates */
  isDate: boolean;
  /** Unique string values (for shared items) */
  sharedItems: string[];
  /** Min numeric value */
  minValue?: number;
  /** Max numeric value */
  maxValue?: number;
}

/**
 * Pivot field axis assignment
 */
export type PivotFieldAxis = 'row' | 'column' | 'filter' | 'value';

/**
 * Configuration for creating a sheet from an array of objects
 */
export interface SheetFromDataConfig<T extends object = Record<string, unknown>> {
  /** Name of the sheet to create */
  name: string;
  /** Array of objects with the same structure */
  data: T[];
  /** Column definitions (optional - defaults to all keys from first object) */
  columns?: ColumnConfig<T>[];
  /** Apply header styling (bold text) (default: true) */
  headerStyle?: boolean;
  /** Starting cell address (default: 'A1') */
  startCell?: string;
}

/**
 * Column configuration for sheet data
 */
export interface ColumnConfig<T = Record<string, unknown>> {
  /** Key from the object to use for this column */
  key: keyof T;
  /** Header text (optional - defaults to key name) */
  header?: string;
  /** Cell style for data cells in this column */
  style?: CellStyle;
}

/**
 * Rich cell value with optional formula and style.
 * Use this when you need to set value, formula, or style for individual cells.
 */
export interface RichCellValue {
  /** Cell value */
  value?: CellValue;
  /** Formula (without leading '=') */
  formula?: string;
  /** Cell style */
  style?: CellStyle;
}

/**
 * Configuration for creating an Excel Table (ListObject)
 */
export interface TableConfig {
  /** Table name (must be unique within the workbook) */
  name: string;
  /** Data range including headers (e.g., "A1:D10") */
  range: string;
  /** First row contains headers (default: true) */
  headerRow?: boolean;
  /** Show total row at the bottom (default: false) */
  totalRow?: boolean;
  /** Table style configuration */
  style?: TableStyleConfig;
}

/**
 * Table style configuration options
 */
export interface TableStyleConfig {
  /** Built-in table style name (e.g., "TableStyleMedium2", "TableStyleLight1") */
  name?: string;
  /** Show banded/alternating row colors (default: true) */
  showRowStripes?: boolean;
  /** Show banded/alternating column colors (default: false) */
  showColumnStripes?: boolean;
  /** Highlight first column with special formatting (default: false) */
  showFirstColumn?: boolean;
  /** Highlight last column with special formatting (default: false) */
  showLastColumn?: boolean;
}

/**
 * Aggregation functions available for table total row
 */
export type TableTotalFunction = 'sum' | 'count' | 'average' | 'min' | 'max' | 'stdDev' | 'var' | 'countNums' | 'none';

/**
 * Configuration for converting a sheet to JSON objects.
 */
export interface SheetToJsonConfig {
  /**
   * Field names to use for each column.
   * If provided, the first row of data starts at row 1 (or startRow).
   * If not provided, the first row is used as field names.
   */
  fields?: string[];

  /**
   * Starting row (0-based). Defaults to 0.
   * If fields are not provided, this row contains the headers.
   * If fields are provided, this is the first data row.
   */
  startRow?: number;

  /**
   * Starting column (0-based). Defaults to 0.
   */
  startCol?: number;

  /**
   * Ending row (0-based, inclusive). Defaults to the last row with data.
   */
  endRow?: number;

  /**
   * Ending column (0-based, inclusive). Defaults to the last column with data.
   */
  endCol?: number;

  /**
   * If true, stop reading when an empty row is encountered. Defaults to true.
   */
  stopOnEmptyRow?: boolean;

  /**
   * How to serialize Date values. Defaults to 'jsDate'.
   */
  dateHandling?: DateHandling;

  /**
   * If true, return formatted text (as displayed in Excel) instead of raw values.
   * All values will be returned as strings. Defaults to false.
   */
  asText?: boolean;
}
