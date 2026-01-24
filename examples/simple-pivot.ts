/**
 * Simple Pivot Table Test
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();
  const sheet = wb.sheet('Sheet1');

  // Simple data: just 3 columns, 4 rows
  sheet.cell('A1').value = 'Category';
  sheet.cell('B1').value = 'Amount';
  sheet.cell('C1').value = 'Count';

  sheet.cell('A2').value = 'A';
  sheet.cell('B2').value = 100;
  sheet.cell('C2').value = 1;

  sheet.cell('A3').value = 'B';
  sheet.cell('B3').value = 200;
  sheet.cell('C3').value = 2;

  sheet.cell('A4').value = 'A';
  sheet.cell('B4').value = 150;
  sheet.cell('C4').value = 1;

  sheet.cell('A5').value = 'B';
  sheet.cell('B5').value = 250;
  sheet.cell('C5').value = 3;

  // Create a simple pivot table
  wb.createPivotTable({
    name: 'SimplePivot',
    source: 'Sheet1!A1:C5',
    target: 'Sheet2!A1',
  })
    .addRowField('Category')
    .addValueField('Amount', 'sum', 'Sum of Amount');

  await wb.toFile('examples/output/simple-pivot.xlsx');
  console.log('Created: examples/output/simple-pivot.xlsx');
}

main().catch(console.error);
