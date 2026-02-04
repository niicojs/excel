/**
 * toJson Example - Converting Excel sheets to JSON with formatting options
 *
 * Run with: bun examples/to-json.ts
 */
import { Workbook } from '../src';

async function main() {
  // Create a sample workbook with various data types
  const wb = Workbook.create();
  const sheet = wb.addSheet('Sales');

  // Headers
  sheet.cell('A1').value = 'Product';
  sheet.cell('B1').value = 'Price';
  sheet.cell('C1').value = 'Quantity';
  sheet.cell('D1').value = 'Total';
  sheet.cell('E1').value = 'Date';
  sheet.cell('F1').value = 'InStock';

  // Style headers
  sheet.range('A1:F1').style = { bold: true, fill: '#4472C4', fontColor: '#FFFFFF' };

  // Data rows with various types
  const products = [
    { product: 'Laptop', price: 1299.99, qty: 5, date: new Date('2024-03-15'), inStock: true },
    { product: 'Mouse', price: 29.5, qty: 50, date: new Date('2024-03-16'), inStock: true },
    { product: 'Keyboard', price: 89.0, qty: 30, date: new Date('2024-03-17'), inStock: false },
    { product: 'Monitor', price: 449.99, qty: 12, date: new Date('2024-03-18'), inStock: true },
    { product: 'Headset', price: 159.0, qty: 0, date: new Date('2024-03-19'), inStock: false },
  ];

  let row = 2;
  for (const p of products) {
    sheet.cell(`A${row}`).value = p.product;
    sheet.cell(`B${row}`).value = p.price;
    sheet.cell(`B${row}`).style = { numberFormat: '$#,##0.00' };
    sheet.cell(`C${row}`).value = p.qty;
    sheet.cell(`D${row}`).formula = `B${row}*C${row}`;
    sheet.cell(`D${row}`).style = { numberFormat: '$#,##0.00' };
    sheet.cell(`E${row}`).value = p.date;
    sheet.cell(`E${row}`).style = { numberFormat: 'yyyy-mm-dd' };
    sheet.cell(`F${row}`).value = p.inStock;
    row++;
  }

  // Save and reload to simulate reading an existing file
  const buffer = await wb.toBuffer();
  const loaded = await Workbook.fromBuffer(buffer);
  const loadedSheet = loaded.sheet('Sales');

  console.log('='.repeat(60));
  console.log('toJson() Example - Raw Values vs Formatted Text');
  console.log('='.repeat(60));

  // Default: raw values
  console.log('\n1. Default toJson() - Returns raw values:');
  console.log('-'.repeat(40));
  const rawData = loadedSheet.toJson();
  console.log('First row:', rawData[0]);
  console.log('Types:', {
    Product: typeof rawData[0].Product,
    Price: typeof rawData[0].Price,
    Quantity: typeof rawData[0].Quantity,
    Date: rawData[0].Date instanceof Date ? 'Date' : typeof rawData[0].Date,
    InStock: typeof rawData[0].InStock,
  });

  // With asText: formatted text
  console.log('\n2. toJson({ asText: true }) - Returns formatted text:');
  console.log('-'.repeat(40));
  const textData = loadedSheet.toJson({ asText: true });
  console.log('First row:', textData[0]);
  console.log('Types: All values are strings');

  // Compare all rows
  console.log('\n3. Full comparison:');
  console.log('-'.repeat(40));
  console.log('\nRaw values:');
  console.table(rawData);

  console.log('\nFormatted text (asText: true):');
  console.table(textData);

  // Practical use case: Export to CSV-like format
  console.log('\n4. Practical use case - CSV-like export:');
  console.log('-'.repeat(40));
  const csvLike = loadedSheet.toJson({ asText: true });
  const headers = Object.keys(csvLike[0]);
  console.log(headers.join('\t'));
  for (const row of csvLike) {
    console.log(Object.values(row).join('\t'));
  }

  // Combined with other options
  console.log('\n5. Combined with other options:');
  console.log('-'.repeat(40));
  const subset = loadedSheet.toJson({
    asText: true,
    startCol: 0,
    endCol: 3, // Only Product, Price, Quantity, Total
  });
  console.log('Subset of columns as text:');
  console.table(subset);
}

main().catch(console.error);
