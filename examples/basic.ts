/**
 * Basic Example - Creating a simple spreadsheet
 *
 * Run with: npx tsx examples/basic.ts
 */
import { Workbook } from '../src';

async function main() {
  // Create a new workbook
  const wb = Workbook.create();
  const sheet = wb.sheet('Sheet1');

  // Write some basic values
  sheet.cell('A1').value = 'Name';
  sheet.cell('B1').value = 'Age';
  sheet.cell('C1').value = 'City';

  // Write data rows
  sheet.cell('A2').value = 'Alice';
  sheet.cell('B2').value = 30;
  sheet.cell('C2').value = 'New York';

  sheet.cell('A3').value = 'Bob';
  sheet.cell('B3').value = 25;
  sheet.cell('C3').value = 'Los Angeles';

  sheet.cell('A4').value = 'Charlie';
  sheet.cell('B4').value = 35;
  sheet.cell('C4').value = 'Chicago';

  // Add a formula
  sheet.cell('A6').value = 'Average Age:';
  sheet.cell('B6').formula = 'AVERAGE(B2:B4)';

  // Save the workbook
  await wb.toFile('examples/output/basic.xlsx');
  console.log('Created: examples/output/basic.xlsx');
}

main().catch(console.error);
