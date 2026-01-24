# Examples

This folder contains example code demonstrating how to use `@niicojs/excel`.

## Setup

First, make sure you have built the library:

```bash
bun run build
```

Create the output directory:

```bash
mkdir -p examples/output
```

Install tsx for running TypeScript directly (if not already installed):

```bash
bun add -d tsx
```

## Running Examples

Each example can be run with:

```bash
npx tsx examples/<example-name>.ts
```

## Examples

### basic.ts
Creates a simple spreadsheet with basic values and a formula.

```bash
npx tsx examples/basic.ts
```

### styling.ts
Demonstrates cell styling including fonts, colors, borders, and number formats.

```bash
npx tsx examples/styling.ts
```

### data-array.ts
Shows how to work with 2D arrays for reading and writing data.

```bash
npx tsx examples/data-array.ts
```

### merged-cells.ts
Creates a report with merged cell headers.

```bash
npx tsx examples/merged-cells.ts
```

### multiple-sheets.ts
Working with multiple worksheets in a workbook.

```bash
npx tsx examples/multiple-sheets.ts
```

### formulas.ts
Various Excel formulas including math, statistics, and conditional logic.

```bash
npx tsx examples/formulas.ts
```

### read-file.ts
Loading and reading data from an existing Excel file.

```bash
npx tsx examples/read-file.ts
```

## Output

All generated Excel files are saved to `examples/output/`. This directory is gitignored.
