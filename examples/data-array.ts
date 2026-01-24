/**
 * Data Array Example - Working with 2D arrays
 *
 * Run with: bun examples/data-array.ts
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();
  const sheet = wb.addSheet('example');

  // Write a 2D array starting at A1
  const salesData = [
    ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'],
    ['North', 1200, 1350, 1100, 1500, null],
    ['South', 980, 1100, 1250, 1180, null],
    ['East', 1500, 1420, 1380, 1600, null],
    ['West', 1100, 1050, 1200, 1350, null],
  ];

  // Write all data at once
  sheet.cell('A1').values = salesData;

  // Add formulas for totals
  for (let row = 2; row <= 5; row++) {
    sheet.cell(`F${row}`).formula = `SUM(B${row}:E${row})`;
  }

  // Add a grand total row
  sheet.cell('A7').value = 'Grand Total';
  for (let col = 1; col <= 5; col++) {
    const colLetter = String.fromCharCode(65 + col); // B, C, D, E, F
    sheet.cell(`${colLetter}7`).formula = `SUM(${colLetter}2:${colLetter}5)`;
  }

  // Style the header row
  sheet.range('A1:F1').style = {
    bold: true,
    fill: '#2E75B6',
    fontColor: '#FFFFFF',
  };

  // Style the total row
  sheet.range('A7:F7').style = {
    bold: true,
    fill: '#BDD7EE',
  };

  // Read the data back as a 2D array
  console.log('Reading data back:');
  const values = sheet.range('A1:F7').values;
  console.table(values);

  await wb.toFile('examples/output/data-array.xlsx');
  console.log('\nCreated: examples/output/data-array.xlsx');
}

main().catch(console.error);
