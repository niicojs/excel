/**
 * Sheet from Objects Example - Creating a sheet from an array of objects
 *
 * Run with: bun examples/sheet-from-objects.ts
 */
import { Workbook } from '../src';
import type { ColumnConfig } from '../src';

async function main() {
  const wb = Workbook.create();

  // Define your data as an array of objects
  interface Employee {
    name: string;
    department: string;
    salary: number;
    startDate: Date;
    active: boolean;
  }

  const employees: Employee[] = [
    {
      name: 'Alice Johnson',
      department: 'Engineering',
      salary: 95000,
      startDate: new Date('2021-03-15'),
      active: true,
    },
    { name: 'Bob Smith', department: 'Marketing', salary: 72000, startDate: new Date('2020-07-01'), active: true },
    {
      name: 'Charlie Brown',
      department: 'Engineering',
      salary: 88000,
      startDate: new Date('2022-01-10'),
      active: true,
    },
    { name: 'Diana Ross', department: 'Sales', salary: 68000, startDate: new Date('2019-11-20'), active: false },
    { name: 'Eve Wilson', department: 'Engineering', salary: 105000, startDate: new Date('2018-05-03'), active: true },
  ];

  // Simple usage - all object keys become columns automatically
  wb.addSheetFromData({
    name: 'All Employees',
    data: employees,
  });

  console.log('Created sheet "All Employees" with all columns from object keys');

  // Custom column configuration - select specific columns with custom headers
  const columns: ColumnConfig<Employee>[] = [
    { key: 'name', header: 'Full Name' },
    { key: 'department', header: 'Dept' },
    { key: 'salary', header: 'Annual Salary', style: { numberFormat: '$#,##0' } },
    { key: 'active', header: 'Currently Active' },
  ];

  wb.addSheetFromData({
    name: 'Custom Columns',
    data: employees,
    columns,
  });

  console.log('Created sheet "Custom Columns" with selected columns and custom headers');

  // Starting at a different position (useful for adding titles above the table)
  const summarySheet = wb.addSheetFromData({
    name: 'With Title',
    data: employees,
    columns: [
      { key: 'name', header: 'Employee' },
      { key: 'department', header: 'Department' },
      { key: 'salary', header: 'Salary' },
    ],
    startCell: 'B3', // Leave room for a title
  });

  // Add a title above the table
  summarySheet.cell('B1').value = 'Employee Summary Report';
  summarySheet.cell('B1').style = { bold: true, fontSize: 16 };
  summarySheet.mergeCells('B1:D1');

  console.log('Created sheet "With Title" starting at B3 with merged title');

  // Without header styling
  wb.addSheetFromData({
    name: 'Plain Headers',
    data: employees,
    headerStyle: false, // Headers will not be bold
  });

  console.log('Created sheet "Plain Headers" without bold headers');

  // Using with different data types
  interface Product {
    sku: string;
    name: string;
    price: number;
    inStock: boolean;
    lastUpdated: Date | null;
  }

  const products: Product[] = [
    { sku: 'PROD-001', name: 'Widget Pro', price: 29.99, inStock: true, lastUpdated: new Date('2024-01-15') },
    { sku: 'PROD-002', name: 'Gadget Basic', price: 14.99, inStock: false, lastUpdated: null },
    { sku: 'PROD-003', name: 'Super Gizmo', price: 49.99, inStock: true, lastUpdated: new Date('2024-02-20') },
  ];

  wb.addSheetFromData({
    name: 'Products',
    data: products,
    columns: [
      { key: 'sku', header: 'SKU' },
      { key: 'name', header: 'Product Name' },
      { key: 'price', header: 'Price', style: { numberFormat: '$#,##0.00' } },
      { key: 'inStock', header: 'In Stock' },
      { key: 'lastUpdated', header: 'Last Updated' },
    ],
  });

  console.log('Created sheet "Products" with mixed data types');

  // Delete the default Sheet1
  wb.deleteSheet('Sheet1');

  // Save the workbook
  await wb.toFile('examples/output/sheet-from-objects.xlsx');
  console.log('\nCreated: examples/output/sheet-from-objects.xlsx');
  console.log(`Total sheets: ${wb.sheetCount}`);
  console.log(`Sheet names: ${wb.sheetNames.join(', ')}`);
}

main().catch(console.error);
