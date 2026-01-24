/**
 * Formulas Example - Working with Excel formulas
 *
 * Run with: npx tsx examples/formulas.ts
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();
  const sheet = wb.sheet('Sheet1');

  // Create a financial calculation sheet
  sheet.cell('A1').value = 'Financial Calculator';
  sheet.cell('A1').style = { bold: true, fontSize: 14 };
  sheet.mergeCells('A1:D1');

  // Input section
  sheet.cell('A3').value = 'Inputs';
  sheet.cell('A3').style = { bold: true, fill: '#D9E2F3' };

  sheet.cell('A4').value = 'Principal';
  sheet.cell('B4').value = 10000;
  sheet.cell('B4').style = { numberFormat: '$#,##0.00' };

  sheet.cell('A5').value = 'Annual Rate';
  sheet.cell('B5').value = 0.05;
  sheet.cell('B5').style = { numberFormat: '0.00%' };

  sheet.cell('A6').value = 'Years';
  sheet.cell('B6').value = 5;

  sheet.cell('A7').value = 'Compounds/Year';
  sheet.cell('B7').value = 12;

  // Calculations section
  sheet.cell('A9').value = 'Calculations';
  sheet.cell('A9').style = { bold: true, fill: '#D9E2F3' };

  // Simple interest
  sheet.cell('A10').value = 'Simple Interest';
  sheet.cell('B10').formula = 'B4*B5*B6';
  sheet.cell('B10').style = { numberFormat: '$#,##0.00' };

  // Compound interest formula: P(1 + r/n)^(nt)
  sheet.cell('A11').value = 'Compound Interest';
  sheet.cell('B11').formula = 'B4*POWER(1+B5/B7,B7*B6)-B4';
  sheet.cell('B11').style = { numberFormat: '$#,##0.00' };

  // Future value
  sheet.cell('A12').value = 'Future Value (Simple)';
  sheet.cell('B12').formula = 'B4+B10';
  sheet.cell('B12').style = { numberFormat: '$#,##0.00' };

  sheet.cell('A13').value = 'Future Value (Compound)';
  sheet.cell('B13').formula = 'B4+B11';
  sheet.cell('B13').style = { numberFormat: '$#,##0.00' };

  // Summary with various functions
  sheet.cell('A15').value = 'Summary Statistics';
  sheet.cell('A15').style = { bold: true, fill: '#D9E2F3' };

  sheet.cell('A16').value = 'Max Future Value';
  sheet.cell('B16').formula = 'MAX(B12:B13)';
  sheet.cell('B16').style = { numberFormat: '$#,##0.00' };

  sheet.cell('A17').value = 'Min Future Value';
  sheet.cell('B17').formula = 'MIN(B12:B13)';
  sheet.cell('B17').style = { numberFormat: '$#,##0.00' };

  sheet.cell('A18').value = 'Difference';
  sheet.cell('B18').formula = 'B13-B12';
  sheet.cell('B18').style = { numberFormat: '$#,##0.00' };

  // Conditional formula
  sheet.cell('A20').value = 'Better Option';
  sheet.cell('B20').formula = 'IF(B13>B12,"Compound","Simple")';

  // Read formulas back
  console.log('Formulas in the sheet:');
  console.log('B10:', sheet.cell('B10').formula);
  console.log('B11:', sheet.cell('B11').formula);
  console.log('B20:', sheet.cell('B20').formula);

  await wb.toFile('examples/output/formulas.xlsx');
  console.log('\nCreated: examples/output/formulas.xlsx');
}

main().catch(console.error);
