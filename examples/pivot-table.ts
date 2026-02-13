/**
 * Pivot Table Example
 *
 * Demonstrates how to create a pivot table with:
 * - Row fields
 * - Value fields
 * - Sorting
 * - Number formatting
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();

  // Source data sheet
  const dataSheet = wb.addSheet('Data');

  dataSheet.cell('A1').value = 'Region';
  dataSheet.cell('B1').value = 'Product';
  dataSheet.cell('C1').value = 'Year';
  dataSheet.cell('D1').value = 'Sales';
  dataSheet.cell('E1').value = 'Quantity';

  const rows = [
    ['North', 'Widget', 2024, 12000, 85],
    ['South', 'Widget', 2024, 9000, 70],
    ['East', 'Gadget', 2024, 15000, 95],
    ['West', 'Gadget', 2024, 11000, 75],
    ['North', 'Widget', 2025, 14000, 90],
    ['South', 'Gadget', 2025, 10500, 78],
    ['East', 'Widget', 2025, 13200, 88],
    ['West', 'Widget', 2025, 12500, 84],
  ];

  for (let index = 0; index < rows.length; index++) {
    const rowNumber = index + 2;
    dataSheet.cell(`A${rowNumber}`).value = rows[index][0];
    dataSheet.cell(`B${rowNumber}`).value = rows[index][1];
    dataSheet.cell(`C${rowNumber}`).value = rows[index][2];
    dataSheet.cell(`D${rowNumber}`).value = rows[index][3];
    dataSheet.cell(`E${rowNumber}`).value = rows[index][4];
  }

  // Target sheet for pivot output
  wb.addSheet('Summary');

  const pivot = wb.createPivotTable({
    name: 'SalesPivot',
    source: 'Data!A1:E9',
    target: 'Summary!A3',
    refreshOnLoad: true,
  });

  pivot
    .addRowField('Region')
    .addRowField('Product')
    .addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00')
    .addValueField('Quantity', 'count', 'Order Count', '0')
    .sortField('Region', 'asc');

  await wb.toFile('examples/output/pivot-table.xlsx');
  console.log('Created: examples/output/pivot-table.xlsx');
}

main().catch(console.error);
