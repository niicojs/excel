/**
 * Pivot Table Example - Creating pivot tables from data
 *
 * Run with: bun examples/big-pivot-table.ts
 */
import { toAddress, Workbook } from '../src';

async function main() {
  const wb = Workbook.create();

  const data = [];
  for (let i = 0; i < 1000; i++) {
    data.push({
      // Date: new Date(2024, Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1),
      Region: ['North', 'South', 'East', 'West'][Math.floor(Math.random() * 4)],
      Product: ['Widget', 'Gadget', 'Doohickey'][Math.floor(Math.random() * 3)],
      Salesperson: ['Alice', 'Bob', 'Charlie', 'Diana'][Math.floor(Math.random() * 4)],
      Units: Math.floor(Math.random() * 20) + 1,
      Sales: parseFloat((Math.random() * 1000).toFixed(2)),
    });
  }

  wb.addSheetFromData({ name: 'SalesData', data });

  console.log(`Created ${data.length} rows of sales data`);

  // Create a pivot table summarizing sales by region and product
  console.log('Creating pivot table: Sales by Region and Product...');

  wb.addSheet('PivotReport1');

  const addr = toAddress(data.length, Object.keys(data[0]).length - 1);

  wb.createPivotTable({
    name: 'SalesByRegionProduct',
    source: `SalesData!A1:${addr}`,
    target: 'PivotReport1!A1',
  })
    .addRowField('Region')
    .addColumnField('Product')
    .addValueField('Sales', 'sum', 'Total Sales')
    .addValueField('Units', 'sum', 'Total Units');

  // Save the workbook
  await wb.toFile('examples/output/big-pivot-table.xlsx');
  console.log('\nCreated: examples/output/big-pivot-table.xlsx');
}

main().catch(console.error);
