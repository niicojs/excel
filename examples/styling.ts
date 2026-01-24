/**
 * Styling Example - Applying styles to cells
 *
 * Run with: npx tsx examples/styling.ts
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();
  const sheet = wb.sheet('Sheet1');

  // Create a styled header row
  sheet.cell('A1').value = 'Product';
  sheet.cell('B1').value = 'Price';
  sheet.cell('C1').value = 'Quantity';
  sheet.cell('D1').value = 'Total';

  // Apply header styles
  sheet.range('A1:D1').style = {
    bold: true,
    fontSize: 12,
    fill: '#4472C4',
    fontColor: '#FFFFFF',
    alignment: { horizontal: 'center' },
    border: {
      bottom: 'medium',
    },
  };

  // Add data with different styles
  const data = [
    ['Widget A', 19.99, 10],
    ['Widget B', 29.99, 5],
    ['Widget C', 9.99, 20],
  ];

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    sheet.cell(`A${row}`).value = data[i][0];
    sheet.cell(`B${row}`).value = data[i][1];
    sheet.cell(`C${row}`).value = data[i][2];
    sheet.cell(`D${row}`).formula = `B${row}*C${row}`;

    // Apply number format to price columns
    sheet.cell(`B${row}`).style = { numberFormat: '$#,##0.00' };
    sheet.cell(`D${row}`).style = { numberFormat: '$#,##0.00' };

    // Alternate row colors
    if (i % 2 === 1) {
      sheet.range(`A${row}:D${row}`).style = { fill: '#D9E2F3' };
    }
  }

  // Add a total row
  const totalRow = data.length + 2;
  sheet.cell(`A${totalRow}`).value = 'TOTAL';
  sheet.cell(`A${totalRow}`).style = { bold: true };
  sheet.cell(`D${totalRow}`).formula = `SUM(D2:D${totalRow - 1})`;
  sheet.cell(`D${totalRow}`).style = {
    bold: true,
    numberFormat: '$#,##0.00',
    border: { top: 'double' },
  };

  await wb.toFile('examples/output/styling.xlsx');
  console.log('Created: examples/output/styling.xlsx');
}

main().catch(console.error);
