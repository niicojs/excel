import { describe, it, expect, beforeEach } from 'vitest';
import { Workbook, Worksheet, Table } from '../src';

describe('Table', () => {
  let wb: Workbook;
  let sheet: Worksheet;

  beforeEach(() => {
    wb = Workbook.create();
    sheet = wb.addSheet('Sheet1');

    // Create sample data
    // Headers
    sheet.cell('A1').value = 'Name';
    sheet.cell('B1').value = 'Department';
    sheet.cell('C1').value = 'Sales';
    sheet.cell('D1').value = 'Quantity';

    // Data rows
    const data = [
      ['Alice', 'North', 1000, 10],
      ['Bob', 'South', 1500, 15],
      ['Charlie', 'East', 1200, 12],
      ['Diana', 'West', 1800, 18],
      ['Eve', 'North', 900, 9],
    ];

    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      sheet.cell(`A${row}`).value = data[i][0];
      sheet.cell(`B${row}`).value = data[i][1];
      sheet.cell(`C${row}`).value = data[i][2];
      sheet.cell(`D${row}`).value = data[i][3];
    }
  });

  describe('createTable', () => {
    it('creates a table with basic configuration', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(table).toBeInstanceOf(Table);
      expect(table.name).toBe('SalesData');
      expect(table.displayName).toBe('SalesData');
      expect(table.range).toBe('A1:D6');
    });

    it('extracts column names from headers', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(table.columns).toEqual(['Name', 'Department', 'Sales', 'Quantity']);
    });

    it('generates default column names when header cells are empty', () => {
      // Create sheet with missing headers
      const sheet2 = wb.addSheet('Sheet2');
      sheet2.cell('A1').value = 'Name';
      // B1 is empty
      sheet2.cell('C1').value = 'Sales';
      sheet2.cell('A2').value = 'Alice';
      sheet2.cell('B2').value = 'North';
      sheet2.cell('C2').value = 100;

      const table = sheet2.createTable({
        name: 'TestTable',
        range: 'A1:C2',
      });

      expect(table.columns).toEqual(['Name', 'Column2', 'Sales']);
    });

    it('adds table to worksheet tables list', () => {
      expect(sheet.tables.length).toBe(0);

      sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(sheet.tables.length).toBe(1);
      expect(sheet.tables[0].name).toBe('SalesData');
    });

    it('throws when table name already exists', () => {
      sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(() => {
        sheet.createTable({
          name: 'SalesData',
          range: 'A1:D6',
        });
      }).toThrow('Table name already exists: SalesData');
    });

    it('throws when table name already exists in another sheet', () => {
      sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const sheet2 = wb.addSheet('Sheet2');
      sheet2.cell('A1').value = 'Test';
      sheet2.cell('A2').value = 'Data';

      expect(() => {
        sheet2.createTable({
          name: 'SalesData',
          range: 'A1:A2',
        });
      }).toThrow('Table name already exists: SalesData');
    });

    it('throws for invalid table name format', () => {
      expect(() => {
        sheet.createTable({
          name: '123Invalid',
          range: 'A1:D6',
        });
      }).toThrow('Invalid table name');

      expect(() => {
        sheet.createTable({
          name: '',
          range: 'A1:D6',
        });
      }).toThrow('Invalid table name');
    });

    it('accepts valid table names with underscores and periods', () => {
      const table = sheet.createTable({
        name: '_MyTable.Data',
        range: 'A1:D6',
      });

      expect(table.name).toBe('_MyTable.Data');
    });
  });

  describe('table properties', () => {
    it('has auto-filter enabled by default', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(table.hasAutoFilter).toBe(true);
    });

    it('does not have total row by default', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(table.hasTotalRow).toBe(false);
    });

    it('creates table with total row when configured', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: true,
      });

      expect(table.hasTotalRow).toBe(true);
    });

    it('has default style configuration', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const style = table.style;
      expect(style.name).toBe('TableStyleMedium2');
      expect(style.showRowStripes).toBe(true);
      expect(style.showColumnStripes).toBe(false);
      expect(style.showFirstColumn).toBe(false);
      expect(style.showLastColumn).toBe(false);
    });

    it('applies custom style configuration', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        style: {
          name: 'TableStyleDark1',
          showRowStripes: false,
          showColumnStripes: true,
          showFirstColumn: true,
          showLastColumn: true,
        },
      });

      const style = table.style;
      expect(style.name).toBe('TableStyleDark1');
      expect(style.showRowStripes).toBe(false);
      expect(style.showColumnStripes).toBe(true);
      expect(style.showFirstColumn).toBe(true);
      expect(style.showLastColumn).toBe(true);
    });
  });

  describe('fluent API', () => {
    it('setAutoFilter returns this for chaining', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const result = table.setAutoFilter(false);
      expect(result).toBe(table);
      expect(table.hasAutoFilter).toBe(false);
    });

    it('setStyle returns this for chaining', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const result = table.setStyle({ name: 'TableStyleLight1' });
      expect(result).toBe(table);
      expect(table.style.name).toBe('TableStyleLight1');
    });

    it('setStyle updates only provided properties', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        style: { showRowStripes: true, showColumnStripes: false },
      });

      table.setStyle({ showColumnStripes: true });

      expect(table.style.showRowStripes).toBe(true);
      expect(table.style.showColumnStripes).toBe(true);
    });

    it('setTotalRow enables/disables total row', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      expect(table.hasTotalRow).toBe(false);

      table.setTotalRow(true);
      expect(table.hasTotalRow).toBe(true);

      table.setTotalRow(false);
      expect(table.hasTotalRow).toBe(false);
    });

    it('supports method chaining', () => {
      const table = sheet
        .createTable({
          name: 'SalesData',
          range: 'A1:D6',
          totalRow: true,
        })
        .setAutoFilter(true)
        .setStyle({ name: 'TableStyleLight5' })
        .setTotalFunction('Sales', 'sum')
        .setTotalFunction('Quantity', 'average');

      expect(table.name).toBe('SalesData');
      expect(table.hasAutoFilter).toBe(true);
      expect(table.style.name).toBe('TableStyleLight5');
    });
  });

  describe('total row functions', () => {
    it('setTotalFunction throws when total row is not enabled', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: false,
      });

      expect(() => {
        table.setTotalFunction('Sales', 'sum');
      }).toThrow('table does not have a total row enabled');
    });

    it('setTotalFunction throws for non-existent column', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: true,
      });

      expect(() => {
        table.setTotalFunction('NonExistent', 'sum');
      }).toThrow('Column not found: NonExistent');
    });

    it('setTotalFunction sets the function for a column', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: true,
      });

      table.setTotalFunction('Sales', 'sum');
      // The function is set internally - verify through XML
      const xml = table.toXml();
      expect(xml).toContain('totalsRowFunction="sum"');
    });

    it('supports all total function types', () => {
      const functions = ['sum', 'count', 'average', 'min', 'max', 'stdDev', 'var', 'countNums'] as const;

      for (const fn of functions) {
        const wb2 = Workbook.create();
        const s = wb2.addSheet('Sheet1');
        s.cell('A1').value = 'Col';
        s.cell('A2').value = 100;

        const table = s.createTable({
          name: `TestTable_${fn}`,
          range: 'A1:A2',
          totalRow: true,
        });

        table.setTotalFunction('Col', fn);
        const xml = table.toXml();
        expect(xml).toContain(`totalsRowFunction="${fn}"`);
      }
    });

    it('writes SUBTOTAL formula to total row cell', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: true,
      });

      table.setTotalFunction('Sales', 'sum');

      // Check that the formula was written to the total row (row 7 for range A1:D7)
      const totalCell = sheet.cell('C7');
      expect(totalCell.formula).toBe('SUBTOTAL(109,[Sales])');
    });
  });

  describe('XML generation', () => {
    it('generates valid table XML', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const xml = table.toXml();

      expect(xml).toContain('<?xml version="1.0"');
      expect(xml).toContain('<table');
      expect(xml).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
      expect(xml).toContain('name="SalesData"');
      expect(xml).toContain('displayName="SalesData"');
      expect(xml).toContain('ref="A1:D6"');
    });

    it('includes autoFilter element when enabled', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const xml = table.toXml();
      expect(xml).toContain('<autoFilter');
      expect(xml).toContain('ref="A1:D6"');
    });

    it('excludes autoFilter when disabled', () => {
      const table = sheet
        .createTable({
          name: 'SalesData',
          range: 'A1:D6',
        })
        .setAutoFilter(false);

      const xml = table.toXml();
      expect(xml).not.toContain('<autoFilter');
    });

    it('includes tableColumns with correct names', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const xml = table.toXml();
      expect(xml).toContain('<tableColumns count="4"');
      expect(xml).toContain('name="Name"');
      expect(xml).toContain('name="Department"');
      expect(xml).toContain('name="Sales"');
      expect(xml).toContain('name="Quantity"');
    });

    it('includes tableStyleInfo with correct attributes', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        style: {
          name: 'TableStyleMedium9',
          showRowStripes: true,
          showColumnStripes: true,
          showFirstColumn: true,
          showLastColumn: false,
        },
      });

      const xml = table.toXml();
      expect(xml).toContain('name="TableStyleMedium9"');
      expect(xml).toContain('showRowStripes="1"');
      expect(xml).toContain('showColumnStripes="1"');
      expect(xml).toContain('showFirstColumn="1"');
      expect(xml).toContain('showLastColumn="0"');
    });

    it('includes totalsRowCount when total row is enabled', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: true,
      });

      const xml = table.toXml();
      expect(xml).toContain('totalsRowCount="1"');
    });

    it('includes totalsRowShown="0" when total row is disabled', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
        totalRow: false,
      });

      const xml = table.toXml();
      expect(xml).toContain('totalsRowShown="0"');
    });

    it('autoFilter excludes total row when present', () => {
      const table = sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6', // Data range (without total row)
        totalRow: true,
      });

      const xml = table.toXml();
      // Table range expands to A1:D7 with total row, autoFilter should end at D6 (excluding total row)
      expect(xml).toContain('ref="A1:D7"'); // Full table range
      expect(xml).toContain('autoFilter ref="A1:D6"'); // AutoFilter excludes total row
    });
  });

  describe('file generation', () => {
    it('generates a valid xlsx file with table', async () => {
      sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);

      // Verify it can be read back
      const wb2 = await Workbook.fromBuffer(buffer);
      expect(wb2.sheetNames).toContain('Sheet1');
    });

    it('generates file with multiple tables in same sheet', async () => {
      sheet.createTable({
        name: 'Table1',
        range: 'A1:B3',
      });

      // Add second data set
      sheet.cell('E1').value = 'Product';
      sheet.cell('F1').value = 'Price';
      sheet.cell('E2').value = 'Widget';
      sheet.cell('F2').value = 100;

      sheet.createTable({
        name: 'Table2',
        range: 'E1:F2',
      });

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);

      const wb2 = await Workbook.fromBuffer(buffer);
      expect(wb2.sheetNames).toContain('Sheet1');
    });

    it('generates file with tables in multiple sheets', async () => {
      sheet.createTable({
        name: 'Table1',
        range: 'A1:D6',
      });

      const sheet2 = wb.addSheet('Sheet2');
      sheet2.cell('A1').value = 'ID';
      sheet2.cell('B1').value = 'Value';
      sheet2.cell('A2').value = 1;
      sheet2.cell('B2').value = 100;

      sheet2.createTable({
        name: 'Table2',
        range: 'A1:B2',
      });

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);

      const wb2 = await Workbook.fromBuffer(buffer);
      expect(wb2.sheetNames).toContain('Sheet1');
      expect(wb2.sheetNames).toContain('Sheet2');
    });

    it('generates file with table having total row', async () => {
      sheet
        .createTable({
          name: 'SalesData',
          range: 'A1:D6',
          totalRow: true,
        })
        .setTotalFunction('Sales', 'sum')
        .setTotalFunction('Quantity', 'average');

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });

    it('generates file with table and pivot table on same sheet', async () => {
      // Create table
      sheet.createTable({
        name: 'SalesData',
        range: 'A1:D6',
      });

      // Create pivot table
      wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D6',
        target: 'Sheet1!F1',
      })
        .addRowField('Department')
        .addValueField('Sales', 'sum');

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });

  describe('worksheet integration', () => {
    it('tables getter returns array copy', () => {
      sheet.createTable({
        name: 'Table1',
        range: 'A1:D6',
      });

      const tables1 = sheet.tables;
      const tables2 = sheet.tables;

      expect(tables1).not.toBe(tables2);
      expect(tables1).toEqual(tables2);
    });

    it('marks worksheet as dirty when table is created', () => {
      const newWb = Workbook.create();
      const newSheet = newWb.addSheet('Test');
      newSheet.cell('A1').value = 'Header';
      newSheet.cell('A2').value = 'Data';

      // Access sheet to reset dirty state (by reading)
      expect(newSheet.dirty).toBe(true);

      newSheet.createTable({
        name: 'TestTable',
        range: 'A1:A2',
      });

      expect(newSheet.dirty).toBe(true);
    });
  });
});
