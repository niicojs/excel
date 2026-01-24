/**
 * Sheet from Objects Example - Creating a sheet from an array of objects
 *
 * Run with: bun examples/sheet-from-objects.ts
 */
import { Workbook } from '../src';
import type { ColumnConfig, RichCellValue } from '../src';

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

  // ====================================================================
  // RichCellValue - Adding formulas and styles to individual cells
  // ====================================================================

  // RichCellValue allows you to set value, formula, and/or style for individual cells
  interface OrderLine {
    product: string;
    unitPrice: number;
    quantity: number;
    total: RichCellValue; // Use RichCellValue for cells with formulas/styles
  }

  const orderLines: OrderLine[] = [
    {
      product: 'Widget Pro',
      unitPrice: 29.99,
      quantity: 10,
      total: { formula: 'B2*C2', style: { numberFormat: '$#,##0.00' } },
    },
    {
      product: 'Gadget Basic',
      unitPrice: 14.99,
      quantity: 25,
      total: { formula: 'B3*C3', style: { numberFormat: '$#,##0.00' } },
    },
    {
      product: 'Super Gizmo',
      unitPrice: 49.99,
      quantity: 5,
      total: { formula: 'B4*C4', style: { numberFormat: '$#,##0.00' } },
    },
  ];

  wb.addSheetFromData({
    name: 'Order with Formulas',
    data: orderLines,
    columns: [
      { key: 'product', header: 'Product' },
      { key: 'unitPrice', header: 'Unit Price', style: { numberFormat: '$#,##0.00' } },
      { key: 'quantity', header: 'Qty' },
      { key: 'total', header: 'Total' },
    ],
  });

  console.log('Created sheet "Order with Formulas" with RichCellValue formulas');

  // RichCellValue with styles for conditional formatting-like effects
  interface StatusReport {
    task: string;
    status: RichCellValue;
    priority: RichCellValue;
  }

  const statusReport: StatusReport[] = [
    {
      task: 'Complete documentation',
      status: { value: 'Done', style: { bold: true } },
      priority: { value: 'High', style: { bold: true, italic: true } },
    },
    {
      task: 'Fix critical bug',
      status: { value: 'In Progress', style: { italic: true } },
      priority: { value: 'Critical', style: { bold: true } },
    },
    {
      task: 'Review PRs',
      status: { value: 'Pending', style: {} },
      priority: { value: 'Medium', style: { italic: true } },
    },
  ];

  wb.addSheetFromData({
    name: 'Status Report',
    data: statusReport,
  });

  console.log('Created sheet "Status Report" with styled status cells');

  // RichCellValue with formulas and grand total
  interface SalesData {
    region: string;
    q1: number;
    q2: number;
    q3: number;
    q4: number;
    yearTotal: RichCellValue;
  }

  const salesData: SalesData[] = [
    {
      region: 'North',
      q1: 10000,
      q2: 12000,
      q3: 11000,
      q4: 15000,
      yearTotal: { formula: 'SUM(B2:E2)', style: { bold: true, numberFormat: '$#,##0' } },
    },
    {
      region: 'South',
      q1: 8000,
      q2: 9500,
      q3: 10500,
      q4: 12000,
      yearTotal: { formula: 'SUM(B3:E3)', style: { bold: true, numberFormat: '$#,##0' } },
    },
    {
      region: 'East',
      q1: 15000,
      q2: 14000,
      q3: 16000,
      q4: 18000,
      yearTotal: { formula: 'SUM(B4:E4)', style: { bold: true, numberFormat: '$#,##0' } },
    },
    {
      region: 'West',
      q1: 12000,
      q2: 13000,
      q3: 11500,
      q4: 14500,
      yearTotal: { formula: 'SUM(B5:E5)', style: { bold: true, numberFormat: '$#,##0' } },
    },
  ];

  wb.addSheetFromData({
    name: 'Sales Summary',
    data: salesData,
    columns: [
      { key: 'region', header: 'Region' },
      { key: 'q1', header: 'Q1', style: { numberFormat: '$#,##0' } },
      { key: 'q2', header: 'Q2', style: { numberFormat: '$#,##0' } },
      { key: 'q3', header: 'Q3', style: { numberFormat: '$#,##0' } },
      { key: 'q4', header: 'Q4', style: { numberFormat: '$#,##0' } },
      { key: 'yearTotal', header: 'Year Total' },
    ],
  });

  console.log('Created sheet "Sales Summary" with SUM formulas and formatting');

  // Save the workbook
  await wb.toFile('examples/output/sheet-from-objects.xlsx');
  console.log('\nCreated: examples/output/sheet-from-objects.xlsx');
  console.log(`Total sheets: ${wb.sheetCount}`);
  console.log(`Sheet names: ${wb.sheetNames.join(', ')}`);
}

main().catch(console.error);
