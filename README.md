# @niicojs/excel

A TypeScript library for Excel/OpenXML manipulation with maximum format preservation.

## Features

- Read and write `.xlsx` files
- Preserve formatting when modifying existing files
- Full formula support (read/write/preserve)
- Cell styles (fonts, fills, borders, alignment)
- Merged cells
- Column widths and row heights
- Freeze panes
- Pivot tables with fluent API
- Sheet operations (add, delete, rename, copy)
- Create sheets from arrays of objects (`addSheetFromData`)
- Convert sheets to JSON arrays (`toJson`)
- Create new workbooks from scratch
- TypeScript-first with full type definitions

## Installation

```bash
pnpm install @niicojs/excel
# or
bun add @niicojs/excel
```

## Quick Start

```typescript
import { Workbook } from '@niicojs/excel';

// Create a new workbook
const wb = Workbook.create();
wb.addSheet('Sheet1');
const sheet = wb.sheet('Sheet1');

// Write data
sheet.cell('A1').value = 'Hello';
sheet.cell('B1').value = 42;
sheet.cell('C1').value = true;
sheet.cell('D1').value = new Date();

// Write formulas
sheet.cell('A2').formula = 'SUM(B1:B1)';

// Write a 2D array
sheet.cell('A3').values = [
  ['Name', 'Age', 'City'],
  ['Alice', 30, 'NYC'],
  ['Bob', 25, 'LA'],
];

// Save to file
await wb.toFile('output.xlsx');
```

## Loading Existing Files

```typescript
import { Workbook } from '@niicojs/excel';

// Load from file
const wb = await Workbook.fromFile('template.xlsx');

// Or load from buffer
const buffer = await fetch('https://example.com/file.xlsx').then((r) => r.arrayBuffer());
const wb = await Workbook.fromBuffer(new Uint8Array(buffer));

// Read data
const sheet = wb.sheet('Sheet1');
console.log(sheet.cell('A1').value); // The cell value
console.log(sheet.cell('A1').formula); // The formula (if any)
console.log(sheet.cell('A1').type); // 'string' | 'number' | 'boolean' | 'date' | 'error' | 'empty'
```

## Working with Ranges

```typescript
const sheet = wb.sheet(0);

// Read a range
const values = sheet.range('A1:C10').values; // 2D array

// Write to a range
sheet.range('A1:B2').values = [
  [1, 2],
  [3, 4],
];

// Get formulas from a range
const formulas = sheet.range('A1:C10').formulas;
```

## Styling Cells

```typescript
const sheet = wb.sheet(0);

// Apply styles to a cell
sheet.cell('A1').style = {
  bold: true,
  italic: true,
  fontSize: 14,
  fontName: 'Arial',
  fontColor: '#FF0000',
  fill: '#FFFF00',
  border: {
    top: 'thin',
    bottom: 'medium',
    left: 'thin',
    right: 'thin',
  },
  alignment: {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  },
  numberFormat: '#,##0.00',
};

// Apply styles to a range
sheet.range('A1:C1').style = { bold: true, fill: '#CCCCCC' };
```

## Merged Cells

```typescript
const sheet = wb.sheet(0);

// Merge cells
sheet.mergeCells('A1:C1');

// Or using two addresses
sheet.mergeCells('A1', 'C1');

// Unmerge
sheet.unmergeCells('A1:C1');

// Get all merged regions
console.log(sheet.mergedCells); // ['A1:C3', 'D5:E10', ...]
```

## Column Widths and Row Heights

```typescript
const sheet = wb.sheet(0);

// Set column width (by letter or 0-based index)
sheet.setColumnWidth('A', 20);
sheet.setColumnWidth(1, 15); // Column B

// Get column width
const width = sheet.getColumnWidth('A'); // 20 or undefined

// Set row height (0-based index)
sheet.setRowHeight(0, 30); // First row

// Get row height
const height = sheet.getRowHeight(0); // 30 or undefined
```

## Freeze Panes

```typescript
const sheet = wb.sheet(0);

// Freeze first row (header)
sheet.freezePane(1, 0);

// Freeze first column
sheet.freezePane(0, 1);

// Freeze first row and first column
sheet.freezePane(1, 1);

// Unfreeze
sheet.freezePane(0, 0);

// Get current freeze pane configuration
const frozen = sheet.getFrozenPane();
// { row: 1, col: 0 } or null
```

## Sheet Operations

```typescript
const wb = Workbook.create();

// Add sheets
wb.addSheet('Data');
wb.addSheet('Summary', 0); // Insert at index 0

// Get sheet names
console.log(wb.sheetNames); // ['Summary', 'Data']

// Access sheets
const sheet = wb.sheet('Data'); // By name
const sheet = wb.sheet(0); // By index

// Rename sheet
wb.renameSheet('Data', 'RawData');

// Copy sheet
wb.copySheet('RawData', 'RawData_Backup');

// Delete sheet
wb.deleteSheet('Summary');
```

## Pivot Tables

Create pivot tables with a fluent API:

```typescript
const wb = await Workbook.fromFile('sales-data.xlsx');

// Create a pivot table from source data
const pivot = wb.createPivotTable({
  name: 'SalesPivot',
  source: 'Data!A1:E100', // Source range with headers
  target: 'Summary!A3', // Where to place the pivot table
  refreshOnLoad: true, // Refresh when file opens (default: true)
});

// Configure fields using fluent API
pivot
  .addRowField('Region') // Group by region
  .addRowField('Product') // Then by product
  .addColumnField('Year') // Columns by year
  .addValueField('Sales', 'sum', 'Total Sales') // Sum of sales
  .addValueField('Quantity', 'count', 'Order Count') // Count of orders
  .addFilterField('Category'); // Page filter

// Sort and filter fields
pivot
  .sortField('Region', 'asc') // Sort ascending
  .filterField('Product', { include: ['Widget', 'Gadget'] }); // Include only these

await wb.toFile('report.xlsx');
```

### Pivot Table API

```typescript
// Add fields to different areas
pivot.addRowField(fieldName: string): PivotTable
pivot.addColumnField(fieldName: string): PivotTable
pivot.addValueField(fieldName: string, aggregation?: AggregationType, displayName?: string): PivotTable
pivot.addFilterField(fieldName: string): PivotTable

// Aggregation types: 'sum' | 'count' | 'average' | 'min' | 'max'

// Sort a row or column field
pivot.sortField(fieldName: string, order: 'asc' | 'desc'): PivotTable

// Filter field values
pivot.filterField(fieldName: string, filter: { include?: string[] } | { exclude?: string[] }): PivotTable
```

## Creating Sheets from Data

Create sheets directly from arrays of objects with `addSheetFromData`:

```typescript
const wb = Workbook.create();

// Simple usage - object keys become column headers
const employees = [
  { name: 'Alice', age: 30, city: 'Paris' },
  { name: 'Bob', age: 25, city: 'London' },
];

wb.addSheetFromData({
  name: 'Employees',
  data: employees,
});

// Custom column configuration
wb.addSheetFromData({
  name: 'Custom',
  data: employees,
  columns: [
    { key: 'name', header: 'Full Name' },
    { key: 'age', header: 'Age (years)' },
    { key: 'city', header: 'Location', style: { bold: true } },
  ],
});

// With formulas and styles using RichCellValue
const orderLines = [
  { product: 'Widget', price: 10, qty: 5, total: { formula: 'B2*C2', style: { bold: true } } },
  { product: 'Gadget', price: 20, qty: 3, total: { formula: 'B3*C3', style: { bold: true } } },
];

wb.addSheetFromData({
  name: 'Orders',
  data: orderLines,
});

// Other options
wb.addSheetFromData({
  name: 'Options',
  data: employees,
  headerStyle: false, // Don't bold headers
  startCell: 'B3', // Start at B3 instead of A1
});
```

## Converting Sheets to JSON

Convert sheet data back to arrays of objects with `toJson`:

```typescript
const sheet = wb.sheet('Data');

// Using first row as headers
const data = sheet.toJson();
// [{ name: 'Alice', age: 30 }, { name: 'Bob', age: 25 }]

// Using custom field names (first row is data, not headers)
const data2 = sheet.toJson({
  fields: ['name', 'age', 'city'],
});

// With TypeScript generics
interface Person {
  name: string | null;
  age: number | null;
}
const people = sheet.toJson<Person>();

// Starting from a specific position
const data3 = sheet.toJson({
  startRow: 2, // Skip first 2 rows (0-based)
  startCol: 1, // Start from column B
});

// Limiting the range
const data4 = sheet.toJson({
  endRow: 10, // Stop at row 11 (0-based, inclusive)
  endCol: 3, // Only read columns A-D
});

// Continue past empty rows
const data5 = sheet.toJson({
  stopOnEmptyRow: false, // Default is true
});

// Control how dates are serialized
const data6 = sheet.toJson({
  dateHandling: 'isoString', // 'jsDate' | 'excelSerial' | 'isoString'
});
```

### Roundtrip Example

```typescript
// Create from objects
const originalData = [
  { name: 'Alice', age: 30 },
  { name: 'Bob', age: 25 },
];

const sheet = wb.addSheetFromData({
  name: 'People',
  data: originalData,
});

// Read back as objects
const readData = sheet.toJson();
// readData equals originalData
```

## Saving

```typescript
// Load from file
const wb = await Workbook.fromFile('template.xlsx');

// Or load from buffer
const buffer = await fetch('https://example.com/file.xlsx').then((r) => r.arrayBuffer());
const wb2 = await Workbook.fromBuffer(new Uint8Array(buffer));

// Read data
const sheet = wb.sheet('Sheet1');
console.log(sheet.cell('A1').value); // The cell value
console.log(sheet.cell('A1').formula); // The formula (if any)
console.log(sheet.cell('A1').type); // 'string' | 'number' | 'boolean' | 'date' | 'error' | 'empty'

// Check if a cell exists without creating it
const existingCell = sheet.getCellIfExists('A1');
if (existingCell) {
  console.log(existingCell.value);
}
```

## Type Definitions

```typescript
// Cell values
type CellValue = number | string | boolean | Date | null | CellError;

interface CellError {
  error: '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A';
}

// Cell types
type CellType = 'number' | 'string' | 'boolean' | 'date' | 'error' | 'empty';

// Border types
type BorderType = 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';

// Configuration for addSheetFromData
interface SheetFromDataConfig<T> {
  name: string;
  data: T[];
  columns?: ColumnConfig<T>[];
  headerStyle?: boolean; // Default: true
  startCell?: string; // Default: 'A1'
}

interface ColumnConfig<T> {
  key: keyof T;
  header?: string;
  style?: CellStyle;
}

// Rich cell value for formulas/styles in data
interface RichCellValue {
  value?: CellValue;
  formula?: string;
  style?: CellStyle;
}

// Configuration for toJson
interface SheetToJsonConfig {
  fields?: string[];
  startRow?: number;
  startCol?: number;
  endRow?: number;
  endCol?: number;
  stopOnEmptyRow?: boolean; // Default: true
  dateHandling?: 'jsDate' | 'excelSerial' | 'isoString'; // Default: 'jsDate'
}

// Pivot table configuration
interface PivotTableConfig {
  name: string;
  source: string; // e.g., "Sheet1!A1:D100"
  target: string; // e.g., "Sheet2!A3"
  refreshOnLoad?: boolean; // Default: true
}

type AggregationType = 'sum' | 'count' | 'average' | 'min' | 'max';
type PivotSortOrder = 'asc' | 'desc';
interface PivotFieldFilter {
  include?: string[];
  exclude?: string[];
}
```

## Style Schema

Supported style properties (CellStyle):

```typescript
interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'single' | 'double';
  strike?: boolean;
  fontSize?: number;
  fontName?: string;
  fontColor?: string; // Hex (RGB/RRGGBB/AARRGGBB)
  fill?: string; // Hex (RGB/RRGGBB/AARRGGBB)
  border?: {
    top?: 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';
    bottom?: 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';
    left?: 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';
    right?: 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'justify';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
    textRotation?: number;
  };
  numberFormat?: string; // Excel format code
}
```

## Performance and Large Files

Tips for large sheets and high-volume writes:

- Prefer `addSheetFromData` for bulk writes when possible.
- Use `range.getValues({ createMissing: false })` to avoid creating empty cells during reads.
- Keep shared strings small: prefer numbers/booleans where applicable.
- Avoid frequent `toBuffer()` calls in loops; batch writes and serialize once.

## Limitations

The library focuses on preserving existing structure and editing common parts of workbooks.

- Chart editing is not supported (charts are preserved only).
- Conditional formatting and data validation are preserved but not editable yet.
- Some advanced Excel features (sparklines, slicers, macros) are preserved only.

## Format Preservation

When modifying existing Excel files, this library preserves:

- Cell formatting and styles
- Formulas
- Charts and images
- Merged cells
- Conditional formatting
- Data validation
- And other Excel features

This is achieved by only modifying what's necessary and keeping the original XML structure intact.

## License

MIT
