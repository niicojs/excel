/**
 * Multiple Sheets Example - Working with multiple worksheets
 *
 * Run with: bun examples/multiple-sheets.ts
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();

  // Rename the default sheet
  wb.renameSheet('Sheet1', 'Summary');

  // Add more sheets
  wb.addSheet('January');
  wb.addSheet('February');
  wb.addSheet('March');

  // Populate January data
  const january = wb.sheet('January');
  january.cell('A1').value = 'January Sales';
  january.cell('A1').style = { bold: true, fontSize: 14 };
  january.cell('A3').values = [
    ['Product', 'Sales'],
    ['Widget A', 1200],
    ['Widget B', 850],
    ['Widget C', 1500],
  ];
  january.cell('A7').value = 'Total';
  january.cell('B7').formula = 'SUM(B4:B6)';

  // Populate February data
  const february = wb.sheet('February');
  february.cell('A1').value = 'February Sales';
  february.cell('A1').style = { bold: true, fontSize: 14 };
  february.cell('A3').values = [
    ['Product', 'Sales'],
    ['Widget A', 1350],
    ['Widget B', 920],
    ['Widget C', 1680],
  ];
  february.cell('A7').value = 'Total';
  february.cell('B7').formula = 'SUM(B4:B6)';

  // Populate March data
  const march = wb.sheet('March');
  march.cell('A1').value = 'March Sales';
  march.cell('A1').style = { bold: true, fontSize: 14 };
  march.cell('A3').values = [
    ['Product', 'Sales'],
    ['Widget A', 1100],
    ['Widget B', 980],
    ['Widget C', 1420],
  ];
  march.cell('A7').value = 'Total';
  march.cell('B7').formula = 'SUM(B4:B6)';

  // Create summary sheet
  const summary = wb.sheet('Summary');
  summary.cell('A1').value = 'Quarterly Summary';
  summary.cell('A1').style = { bold: true, fontSize: 16 };
  summary.mergeCells('A1:C1');

  summary.cell('A3').values = [
    ['Month', 'Total Sales'],
    ['January', null],
    ['February', null],
    ['March', null],
    ['Q1 Total', null],
  ];

  // Reference other sheets (formulas would reference other sheets in real Excel)
  summary.cell('B4').value = 3550; // January total
  summary.cell('B5').value = 3950; // February total
  summary.cell('B6').value = 3500; // March total
  summary.cell('B7').formula = 'SUM(B4:B6)';

  summary.range('A3:B3').style = { bold: true, fill: '#D9E2F3' };
  summary.range('A7:B7').style = { bold: true, fill: '#BDD7EE' };

  // Copy a sheet
  wb.copySheet('January', 'January_Backup');

  // List all sheets
  console.log('Sheets in workbook:', wb.sheetNames);
  console.log('Sheet count:', wb.sheetCount);

  await wb.toFile('examples/output/multiple-sheets.xlsx');
  console.log('Created: examples/output/multiple-sheets.xlsx');
}

main().catch(console.error);
