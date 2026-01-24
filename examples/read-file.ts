/**
 * Read File Example - Loading and reading an existing Excel file
 *
 * Run with: npx tsx examples/read-file.ts
 */
import { Workbook } from '../src';

async function main() {
  // First, create a sample file to read
  const wb = Workbook.create();
  const sheet = wb.sheet('Sheet1');

  sheet.cell('A1').value = 'Employee Data';
  sheet.mergeCells('A1:D1');
  sheet.cell('A1').style = { bold: true, fontSize: 14 };

  sheet.cell('A3').values = [
    ['ID', 'Name', 'Department', 'Salary'],
    [1, 'John Smith', 'Engineering', 75000],
    [2, 'Jane Doe', 'Marketing', 65000],
    [3, 'Bob Wilson', 'Sales', 70000],
    [4, 'Alice Brown', 'Engineering', 80000],
  ];

  sheet.cell('A8').value = 'Average Salary:';
  sheet.cell('B8').formula = 'AVERAGE(D4:D7)';
  sheet.cell('B8').style = { numberFormat: '$#,##0.00' };

  await wb.toFile('examples/output/sample.xlsx');
  console.log('Created sample file: examples/output/sample.xlsx\n');

  // Now read it back
  console.log('Reading the file back...\n');
  const loaded = await Workbook.fromFile('examples/output/sample.xlsx');

  // List sheets
  console.log('Sheets:', loaded.sheetNames);

  // Read the sheet
  const loadedSheet = loaded.sheet(0);
  console.log('Sheet name:', loadedSheet.name);

  // Read individual cells
  console.log('\nCell A1 value:', loadedSheet.cell('A1').value);
  console.log('Cell A1 type:', loadedSheet.cell('A1').type);

  // Read a range
  console.log('\nData range (A3:D7):');
  const data = loadedSheet.range('A3:D7').values;
  console.table(data);

  // Read formulas
  console.log('\nFormula in B8:', loadedSheet.cell('B8').formula);

  // Check merged cells
  console.log('\nMerged regions:', loadedSheet.mergedCells);

  // Iterate over cells
  console.log('\nAll cells with values:');
  for (const [address, cell] of loadedSheet.cells) {
    if (cell.value !== null) {
      console.log(`  ${address}: ${cell.value} (${cell.type})`);
    }
  }
}

main().catch(console.error);
