/**
 * Excel Table Example
 *
 * Demonstrates how to create Excel Tables (ListObjects) with:
 * - Auto-filter
 * - Banded styling
 * - Total row with aggregation functions
 */
import { Workbook } from '../src';

async function main() {
  const wb = Workbook.create();
  const sheet = wb.addSheet('Sales Report');

  // Create headers
  sheet.cell('A1').value = 'Region';
  sheet.cell('B1').value = 'Product';
  sheet.cell('C1').value = 'Sales';
  sheet.cell('D1').value = 'Quantity';
  sheet.cell('E1').value = 'Unit Price';

  // Add sales data
  const data = [
    ['North', 'Widget', 15000, 100, 150],
    ['South', 'Gadget', 22000, 110, 200],
    ['East', 'Widget', 18500, 125, 148],
    ['West', 'Gadget', 24000, 120, 200],
    ['North', 'Gadget', 19500, 95, 205],
    ['South', 'Widget', 16000, 105, 152],
    ['East', 'Gadget', 21000, 100, 210],
    ['West', 'Widget', 17500, 115, 152],
  ];

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    sheet.cell(`A${row}`).value = data[i][0];
    sheet.cell(`B${row}`).value = data[i][1];
    sheet.cell(`C${row}`).value = data[i][2];
    sheet.cell(`D${row}`).value = data[i][3];
    sheet.cell(`E${row}`).value = data[i][4];
  }

  // Create a table with total row
  const table = sheet.createTable({
    name: 'SalesData',
    range: 'A1:E9', // Headers + 8 data rows
    totalRow: true,
    style: {
      name: 'TableStyleMedium9',
      showRowStripes: true,
      showColumnStripes: false,
    },
  });

  // Add total functions to numeric columns
  table
    .setTotalFunction('Sales', 'sum')
    .setTotalFunction('Quantity', 'sum')
    .setTotalFunction('Unit Price', 'average');

  console.log('Table created:', table.name);
  console.log('Table range:', table.range);
  console.log('Columns:', table.columns);

  // Create a second table on another sheet
  const sheet2 = wb.addSheet('Inventory');

  sheet2.cell('A1').value = 'Item';
  sheet2.cell('B1').value = 'Stock';
  sheet2.cell('C1').value = 'Reorder Level';

  const inventory = [
    ['Widget', 500, 100],
    ['Gadget', 350, 75],
    ['Sprocket', 200, 50],
    ['Cog', 150, 40],
  ];

  for (let i = 0; i < inventory.length; i++) {
    const row = i + 2;
    sheet2.cell(`A${row}`).value = inventory[i][0];
    sheet2.cell(`B${row}`).value = inventory[i][1];
    sheet2.cell(`C${row}`).value = inventory[i][2];
  }

  // Create a simple table without total row
  sheet2.createTable({
    name: 'InventoryTable',
    range: 'A1:C5',
    style: { name: 'TableStyleLight14' },
  });

  await wb.toFile('examples/output/table.xlsx');
  console.log('Created: examples/output/table.xlsx');
}

main().catch(console.error);
