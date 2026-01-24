/**
 * Merged Cells Example - Creating headers with merged cells
 *
 * Run with: bun examples/merged-cells.ts
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();
  const sheet = wb.sheet('Sheet1');

  // Create a report title spanning multiple columns
  sheet.cell('A1').value = 'Monthly Sales Report - 2024';
  sheet.mergeCells('A1:E1');
  sheet.cell('A1').style = {
    bold: true,
    fontSize: 16,
    alignment: { horizontal: 'center' },
    fill: '#1F4E79',
    fontColor: '#FFFFFF',
  };

  // Create category headers
  sheet.cell('A3').value = 'Product Info';
  sheet.mergeCells('A3:B3');
  sheet.cell('A3').style = {
    bold: true,
    fill: '#5B9BD5',
    fontColor: '#FFFFFF',
    alignment: { horizontal: 'center' },
  };

  sheet.cell('C3').value = 'Sales Data';
  sheet.mergeCells('C3:E3');
  sheet.cell('C3').style = {
    bold: true,
    fill: '#5B9BD5',
    fontColor: '#FFFFFF',
    alignment: { horizontal: 'center' },
  };

  // Sub-headers
  sheet.cell('A4').value = 'Code';
  sheet.cell('B4').value = 'Name';
  sheet.cell('C4').value = 'Units';
  sheet.cell('D4').value = 'Price';
  sheet.cell('E4').value = 'Revenue';
  sheet.range('A4:E4').style = {
    bold: true,
    fill: '#BDD7EE',
  };

  // Data
  const products = [
    ['P001', 'Widget Pro', 150, 29.99],
    ['P002', 'Widget Basic', 320, 14.99],
    ['P003', 'Widget Deluxe', 85, 49.99],
  ];

  for (let i = 0; i < products.length; i++) {
    const row = 5 + i;
    sheet.cell(`A${row}`).value = products[i][0];
    sheet.cell(`B${row}`).value = products[i][1];
    sheet.cell(`C${row}`).value = products[i][2];
    sheet.cell(`D${row}`).value = products[i][3];
    sheet.cell(`E${row}`).formula = `C${row}*D${row}`;
  }

  // List all merged regions
  console.log('Merged cell regions:', sheet.mergedCells);

  await wb.toFile('examples/output/merged-cells.xlsx');
  console.log('Created: examples/output/merged-cells.xlsx');
}

main().catch(console.error);
