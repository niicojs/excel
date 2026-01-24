# @niicojs/excel

A TypeScript library for Excel/OpenXML manipulation with maximum format preservation.

## Features

- Read and write `.xlsx` files
- Preserve formatting when modifying existing files
- Full formula support (read/write/preserve)
- Cell styles (fonts, fills, borders, alignment)
- Merged cells
- Sheet operations (add, delete, rename, copy)
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

## Sheet Operations

```typescript
const wb = Workbook.create();

// Add sheets
wb.addSheet('Data');
wb.addSheet('Summary', 0); // Insert at index 0

// Get sheet names
console.log(wb.sheetNames); // ['Summary', 'Sheet1', 'Data']

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

## Saving

```typescript
// Save to file
await wb.toFile('output.xlsx');

// Save to buffer (Uint8Array)
const buffer = await wb.toBuffer();
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
```

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
