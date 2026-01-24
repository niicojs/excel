/**
 * Pivot Table Example - Creating pivot tables from data
 *
 * Run with: bun examples/pivot-table.ts
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();

  wb.addSheet('SalesData');
  const sheet = wb.sheet('SalesData');

  // Create sample sales data
  // Headers
  sheet.cell('A1').value = 'Region';
  sheet.cell('B1').value = 'Product';
  sheet.cell('C1').value = 'Salesperson';
  sheet.cell('D1').value = 'Quarter';
  sheet.cell('E1').value = 'Sales';
  sheet.cell('F1').value = 'Units';

  // Style headers
  sheet.range('A1:F1').style = {
    bold: true,
    fill: '#4472C4',
    fontColor: '#FFFFFF',
  };

  // Sample data
  const regions = ['North', 'South', 'East', 'West'];
  const products = ['Laptop', 'Desktop', 'Tablet', 'Phone'];
  const salespeople = ['Alice', 'Bob', 'Carol', 'Dave'];
  const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];

  let row = 2;
  for (const region of regions) {
    for (const product of products) {
      for (const quarter of quarters) {
        const salesperson = salespeople[Math.floor(Math.random() * salespeople.length)];
        const sales = Math.floor(Math.random() * 50000) + 10000;
        const units = Math.floor(Math.random() * 100) + 10;

        sheet.cell(`A${row}`).value = region;
        sheet.cell(`B${row}`).value = product;
        sheet.cell(`C${row}`).value = salesperson;
        sheet.cell(`D${row}`).value = quarter;
        sheet.cell(`E${row}`).value = sales;
        sheet.cell(`F${row}`).value = units;
        row++;
      }
    }
  }

  const lastDataRow = row - 1;
  console.log(`Created ${lastDataRow - 1} rows of sales data`);

  // Create a pivot table summarizing sales by region and product
  console.log('Creating pivot table: Sales by Region and Product...');

  wb.addSheet('PivotReport1');

  wb.createPivotTable({
    name: 'SalesByRegionProduct',
    source: `SalesData!A1:F${lastDataRow}`,
    target: 'PivotReport1!A3',
  })
    .addRowField('Region')
    .addColumnField('Product')
    .addValueField('Sales', 'sum', 'Total Sales')
    .addValueField('Units', 'sum', 'Total Units');

  // Create another pivot table showing sales by quarter
  console.log('Creating pivot table: Sales by Quarter...');

  wb.addSheet('PivotReport2');

  wb.createPivotTable({
    name: 'SalesByQuarter',
    source: `SalesData!A1:F${lastDataRow}`,
    target: 'PivotReport2!A20',
  })
    .addRowField('Quarter')
    .addValueField('Sales', 'sum', 'Total Sales')
    .addValueField('Sales', 'average', 'Avg Sale')
    .addValueField('Sales', 'count', 'Count')
    .addFilterField('Region');

  // Create a pivot table with salesperson performance
  console.log('Creating pivot table: Salesperson Performance...');

  wb.addSheet('PivotReport3');

  wb.createPivotTable({
    name: 'SalespersonPerformance',
    source: `SalesData!A1:F${lastDataRow}`,
    target: 'PivotReport3!H3',
  })
    .addRowField('Salesperson')
    .addValueField('Sales', 'sum', 'Total Sales')
    .addValueField('Sales', 'average', 'Avg Sale')
    .addValueField('Units', 'sum', 'Total Units')
    .addValueField('Sales', 'max', 'Best Sale')
    .addValueField('Sales', 'min', 'Worst Sale');

  // Save the workbook
  await wb.toFile('examples/output/pivot-table.xlsx');
  console.log('\nCreated: examples/output/pivot-table.xlsx');
}

main().catch(console.error);
