import { describe, it, expect, beforeAll } from 'vitest';
import { Workbook } from '../src';
import type { ColumnConfig } from '../src';
import { unlink, mkdir } from 'fs/promises';
import { existsSync } from 'fs';

describe('Workbook', () => {
  const testDir = 'test/fixtures';
  const testFile = `${testDir}/test-output.xlsx`;

  beforeAll(async () => {
    if (!existsSync(testDir)) {
      await mkdir(testDir, { recursive: true });
    }
  });

  describe('create', () => {
    it('creates a new empty workbook', () => {
      const wb = Workbook.create();
      expect(wb.sheetCount).toBe(0);
      expect(wb.sheetNames).toEqual([]);
    });

    it('allows adding sheets to an empty workbook', () => {
      const wb = Workbook.create();
      wb.addSheet('MySheet');
      expect(wb.sheetCount).toBe(1);
      expect(wb.sheetNames).toEqual(['MySheet']);
    });
  });

  describe('sheet operations', () => {
    it('adds a new sheet', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      wb.addSheet('Sheet2');
      expect(wb.sheetCount).toBe(2);
      expect(wb.sheetNames).toEqual(['Sheet1', 'Sheet2']);
    });

    it('adds a sheet at a specific index', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      wb.addSheet('First', 0);
      expect(wb.sheetNames).toEqual(['First', 'Sheet1']);
    });

    it('throws when adding a duplicate sheet name', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      expect(() => wb.addSheet('Sheet1')).toThrow('Sheet already exists');
    });

    it('deletes a sheet by name', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      wb.addSheet('Sheet2');
      wb.deleteSheet('Sheet1');
      expect(wb.sheetNames).toEqual(['Sheet2']);
    });

    it('deletes a sheet by index', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      wb.addSheet('Sheet2');
      wb.deleteSheet(0);
      expect(wb.sheetNames).toEqual(['Sheet2']);
    });

    it('throws when deleting the last sheet', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      expect(() => wb.deleteSheet(0)).toThrow('Cannot delete the last sheet');
    });

    it('renames a sheet', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      wb.renameSheet('Sheet1', 'Renamed');
      expect(wb.sheetNames).toEqual(['Renamed']);
    });

    it('copies a sheet', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const original = wb.sheet('Sheet1');
      original.cell('A1').value = 'Hello';
      original.cell('B1').value = 42;

      const copy = wb.copySheet('Sheet1', 'Copy');
      expect(copy.cell('A1').value).toBe('Hello');
      expect(copy.cell('B1').value).toBe(42);
    });

    it('supports column widths, row heights, and frozen panes', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet('Sheet1');

      sheet.setColumnWidth('B', 22);
      sheet.setRowHeight(2, 18);
      sheet.freezePane(1, 1);

      const frozen = sheet.getFrozenPane();
      expect(frozen).toEqual({ row: 1, col: 1 });
      expect(sheet.getColumnWidth('B')).toBe(22);
      expect(sheet.getRowHeight(2)).toBe(18);
    });

    it('uses lazy range reads to avoid extra cells', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet('Sheet1');

      sheet.cell('A1').value = 1;

      const range = sheet.range('A1:C3');
      const before = sheet.cells.size;
      const values = range.getValues({ createMissing: false });
      const after = sheet.cells.size;

      expect(values).toEqual([
        [1, null, null],
        [null, null, null],
        [null, null, null],
      ]);
      expect(after).toBe(before);
    });
  });

  describe('cell operations', () => {
    it('reads and writes string values', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 'Hello';
      expect(sheet.cell('A1').value).toBe('Hello');
      expect(sheet.cell('A1').type).toBe('string');
    });

    it('reads and writes number values', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 42.5;
      expect(sheet.cell('A1').value).toBe(42.5);
      expect(sheet.cell('A1').type).toBe('number');
    });

    it('reads and writes boolean values', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = true;
      expect(sheet.cell('A1').value).toBe(true);
      expect(sheet.cell('A1').type).toBe('boolean');

      sheet.cell('A2').value = false;
      expect(sheet.cell('A2').value).toBe(false);
    });

    it('reads and writes date values', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);
      const date = new Date('2024-01-15');

      sheet.cell('A1').value = date;
      const result = sheet.cell('A1').value as Date;
      expect(result instanceof Date).toBe(true);
      // Check date components (time may vary due to timezone)
      expect(result.getFullYear()).toBe(2024);
      expect(result.getMonth()).toBe(0); // January
      expect(result.getDate()).toBe(15);
    });

    it('handles error values', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = { error: '#DIV/0!' };

      expect(sheet.cell('A1').type).toBe('error');
      expect(sheet.cell('A1').value).toEqual({ error: '#DIV/0!' });
    });

    it('reads and writes formulas', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 10;
      sheet.cell('A2').value = 20;
      sheet.cell('A3').formula = 'SUM(A1:A2)';

      expect(sheet.cell('A3').formula).toBe('SUM(A1:A2)');
    });

    it('handles null values', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      expect(sheet.cell('A1').value).toBeNull();
      expect(sheet.cell('A1').type).toBe('empty');
    });

    it('handles cell addressing by row/col', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell(0, 0).value = 'A1';
      sheet.cell(1, 2).value = 'C2';

      expect(sheet.cell('A1').value).toBe('A1');
      expect(sheet.cell('C2').value).toBe('C2');
    });
  });

  describe('range operations', () => {
    it('reads values from a range', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 1;
      sheet.cell('B1').value = 2;
      sheet.cell('A2').value = 3;
      sheet.cell('B2').value = 4;

      const values = sheet.range('A1:B2').values;
      expect(values).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    it('writes values to a range', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.range('A1:B2').values = [
        [1, 2],
        [3, 4],
      ];

      expect(sheet.cell('A1').value).toBe(1);
      expect(sheet.cell('B1').value).toBe(2);
      expect(sheet.cell('A2').value).toBe(3);
      expect(sheet.cell('B2').value).toBe(4);
    });

    it('reads range values without creating missing cells', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 10;

      const range = sheet.range('A1:B2');
      const before = sheet.cells.size;
      const values = range.getValues({ createMissing: false });
      const after = sheet.cells.size;

      expect(values).toEqual([
        [10, null],
        [null, null],
      ]);
      expect(after).toBe(before);
    });

    it('writes a 2D array starting at a cell', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('B2').values = [
        ['Name', 'Age'],
        ['Alice', 30],
        ['Bob', 25],
      ];

      expect(sheet.cell('B2').value).toBe('Name');
      expect(sheet.cell('C2').value).toBe('Age');
      expect(sheet.cell('B3').value).toBe('Alice');
      expect(sheet.cell('C3').value).toBe(30);
      expect(sheet.cell('B4').value).toBe('Bob');
      expect(sheet.cell('C4').value).toBe(25);
    });
  });

  describe('merged cells', () => {
    it('merges cells', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.mergeCells('A1:C1');
      expect(sheet.mergedCells).toContain('A1:C1');
    });

    it('merges cells with two arguments', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.mergeCells('A1', 'C1');
      expect(sheet.mergedCells).toContain('A1:C1');
    });

    it('unmerges cells', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.mergeCells('A1:C1');
      sheet.unmergeCells('A1:C1');
      expect(sheet.mergedCells).not.toContain('A1:C1');
    });
  });

  describe('styles', () => {
    it('applies bold style', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').style = { bold: true };
      expect(sheet.cell('A1').style.bold).toBe(true);
    });

    it('applies multiple style properties', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').style = {
        bold: true,
        italic: true,
        fontSize: 14,
      };

      const style = sheet.cell('A1').style;
      expect(style.bold).toBe(true);
      expect(style.italic).toBe(true);
      expect(style.fontSize).toBe(14);
    });

    it('applies style to a range', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.range('A1:B2').style = { bold: true };

      expect(sheet.cell('A1').style.bold).toBe(true);
      expect(sheet.cell('B1').style.bold).toBe(true);
      expect(sheet.cell('A2').style.bold).toBe(true);
      expect(sheet.cell('B2').style.bold).toBe(true);
    });

    it('caches style objects', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const styles = wb.styles;

      const index = styles.createStyle({ bold: true, italic: true });
      const first = styles.getStyle(index);
      const second = styles.getStyle(index);

      expect(first).toEqual(second);
    });

    it('clones styles with overrides', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const styles = wb.styles;

      const baseIndex = styles.createStyle({ bold: true, fontSize: 12 });
      const cloneIndex = styles.cloneStyle(baseIndex, { italic: true });

      const baseStyle = styles.getStyle(baseIndex);
      const cloneStyle = styles.getStyle(cloneIndex);

      expect(baseStyle.bold).toBe(true);
      expect(baseStyle.italic).toBeUndefined();
      expect(cloneStyle.bold).toBe(true);
      expect(cloneStyle.italic).toBe(true);
      expect(cloneStyle.fontSize).toBe(12);
    });
  });

  describe('save and load', () => {
    it('saves and loads a workbook', async () => {
      // Create and populate workbook
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);
      sheet.cell('A1').value = 'Hello';
      sheet.cell('B1').value = 42;
      sheet.cell('C1').value = true;
      sheet.cell('A2').formula = 'SUM(B1:B1)';

      // Save to buffer
      const buffer = await wb.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      // Load from buffer
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      expect(loadedSheet.cell('A1').value).toBe('Hello');
      expect(loadedSheet.cell('B1').value).toBe(42);
      expect(loadedSheet.cell('C1').value).toBe(true);
      expect(loadedSheet.cell('A2').formula).toBe('SUM(B1:B1)');
    });

    it('saves and loads to file', async () => {
      // Create and populate workbook
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);
      sheet.cell('A1').value = 'Test';
      sheet.cell('B1').value = 123;

      // Save to file
      await wb.toFile(testFile);

      // Load from file
      const loaded = await Workbook.fromFile(testFile);
      const loadedSheet = loaded.sheet(0);

      expect(loadedSheet.cell('A1').value).toBe('Test');
      expect(loadedSheet.cell('B1').value).toBe(123);

      // Cleanup
      await unlink(testFile);
    });

    it('preserves merged cells after save/load', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);
      sheet.cell('A1').value = 'Merged Header';
      sheet.mergeCells('A1:C1');

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      expect(loadedSheet.mergedCells).toContain('A1:C1');
    });

    it('preserves styles after save/load', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);
      sheet.cell('A1').value = 'Styled';
      sheet.cell('A1').style = { bold: true, italic: true };

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      const style = loadedSheet.cell('A1').style;
      expect(style.bold).toBe(true);
      expect(style.italic).toBe(true);
    });

    it('preserves column widths, row heights, and freeze panes after save/load', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.setColumnWidth('B', 24);
      sheet.setRowHeight(1, 20);
      sheet.freezePane(1, 2);

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      expect(loadedSheet.getColumnWidth('B')).toBe(24);
      expect(loadedSheet.getRowHeight(1)).toBe(20);
      expect(loadedSheet.getFrozenPane()).toEqual({ row: 1, col: 2 });
    });

    it('preserves merged cells, formulas, and styles together', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 'Header';
      sheet.cell('A1').style = { bold: true, fill: '#FFEEAA' };
      sheet.mergeCells('A1:C1');

      sheet.cell('A2').value = 10;
      sheet.cell('A3').value = 20;
      sheet.cell('A4').formula = 'SUM(A2:A3)';
      sheet.cell('A4').style = { numberFormat: '0.00' };

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      expect(loadedSheet.mergedCells).toContain('A1:C1');
      expect(loadedSheet.cell('A4').formula).toBe('SUM(A2:A3)');
      expect(loadedSheet.cell('A1').style.bold).toBe(true);
      expect(loadedSheet.cell('A4').style.numberFormat).toBe('0.00');
    });

    it('preserves date formats after save/load', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);
      const date = new Date('2024-02-03T15:30:00Z');

      sheet.cell('B2').style = { numberFormat: 'mm/dd/yyyy hh:mm' };
      sheet.cell('B2').value = date;

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      const loadedValue = loadedSheet.cell('B2').value as Date;
      expect(loadedValue).toBeInstanceOf(Date);
      expect(loadedSheet.cell('B2').type).toBe('date');
      expect(loadedSheet.cell('B2').style.numberFormat).toBe('mm/dd/yyyy hh:mm');
    });

    it('preserves error values after save/load', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('C3').value = { error: '#VALUE!' };

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      expect(loadedSheet.cell('C3').type).toBe('error');
      expect(loadedSheet.cell('C3').value).toEqual({ error: '#VALUE!' });
    });

    it('preserves theme-based fill colors after save/load', async () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet1');
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 'Theme';
      sheet.cell('A1').style = { fillTheme: 2, fillTint: 0.5 };

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet(0);

      const style = loadedSheet.cell('A1').style;
      expect(style.fillTheme).toBe(2);
      expect(style.fillTint).toBe(0.5);
    });
  });

  describe('addSheetFromData', () => {
    interface Person {
      name: string;
      age: number;
      city: string;
    }

    const testData: Person[] = [
      { name: 'Alice', age: 30, city: 'Paris' },
      { name: 'Bob', age: 25, city: 'London' },
      { name: 'Charlie', age: 35, city: 'Berlin' },
    ];

    it('creates a sheet from an array of objects', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
      });

      expect(sheet.name).toBe('People');
      expect(wb.sheetNames).toContain('People');
    });

    it('writes headers from object keys', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
      });

      expect(sheet.cell('A1').value).toBe('name');
      expect(sheet.cell('B1').value).toBe('age');
      expect(sheet.cell('C1').value).toBe('city');
    });

    it('writes data values correctly', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
      });

      // First data row (row 2)
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe(30);
      expect(sheet.cell('C2').value).toBe('Paris');

      // Second data row (row 3)
      expect(sheet.cell('A3').value).toBe('Bob');
      expect(sheet.cell('B3').value).toBe(25);
      expect(sheet.cell('C3').value).toBe('London');

      // Third data row (row 4)
      expect(sheet.cell('A4').value).toBe('Charlie');
      expect(sheet.cell('B4').value).toBe(35);
      expect(sheet.cell('C4').value).toBe('Berlin');
    });

    it('applies bold style to headers by default', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
      });

      expect(sheet.cell('A1').style.bold).toBe(true);
      expect(sheet.cell('B1').style.bold).toBe(true);
      expect(sheet.cell('C1').style.bold).toBe(true);
    });

    it('disables header style when headerStyle is false', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
        headerStyle: false,
      });

      expect(sheet.cell('A1').style.bold).toBeUndefined();
    });

    it('uses custom column configuration', () => {
      const wb = Workbook.create();
      const columns: ColumnConfig<Person>[] = [
        { key: 'name', header: 'Full Name' },
        { key: 'age', header: 'Age (years)' },
      ];

      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
        columns,
      });

      // Check custom headers
      expect(sheet.cell('A1').value).toBe('Full Name');
      expect(sheet.cell('B1').value).toBe('Age (years)');

      // City column should not exist (not in columns config)
      expect(sheet.cell('C1').value).toBeNull();

      // Data should still be correct
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe(30);
    });

    it('applies column styles to data cells', () => {
      const wb = Workbook.create();
      const columns: ColumnConfig<Person>[] = [{ key: 'name' }, { key: 'age', style: { italic: true } }];

      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
        columns,
      });

      // Age column data cells should be italic
      expect(sheet.cell('B2').style.italic).toBe(true);
      expect(sheet.cell('B3').style.italic).toBe(true);
      expect(sheet.cell('B4').style.italic).toBe(true);

      // Name column should not have italic
      expect(sheet.cell('A2').style.italic).toBeUndefined();
    });

    it('starts at custom cell position', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'People',
        data: testData,
        startCell: 'C3',
      });

      // Headers at C3, D3, E3
      expect(sheet.cell('C3').value).toBe('name');
      expect(sheet.cell('D3').value).toBe('age');
      expect(sheet.cell('E3').value).toBe('city');

      // First data row at C4
      expect(sheet.cell('C4').value).toBe('Alice');
      expect(sheet.cell('D4').value).toBe(30);
      expect(sheet.cell('E4').value).toBe('Paris');
    });

    it('handles empty data array', () => {
      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'Empty',
        data: [] as Person[],
      });

      expect(sheet.name).toBe('Empty');
      expect(sheet.cell('A1').value).toBeNull();
    });

    it('handles boolean values', () => {
      interface Item {
        name: string;
        active: boolean;
      }

      const data: Item[] = [
        { name: 'Item1', active: true },
        { name: 'Item2', active: false },
      ];

      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'Items',
        data,
      });

      expect(sheet.cell('B2').value).toBe(true);
      expect(sheet.cell('B3').value).toBe(false);
    });

    it('handles date values', () => {
      interface Event {
        name: string;
        date: Date;
      }

      const eventDate = new Date('2024-06-15');
      const data: Event[] = [{ name: 'Conference', date: eventDate }];

      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'Events',
        data,
      });

      const result = sheet.cell('B2').value as Date;
      expect(result instanceof Date).toBe(true);
      expect(result.getFullYear()).toBe(2024);
      expect(result.getMonth()).toBe(5); // June (0-based)
      expect(result.getDate()).toBe(15);
    });

    it('handles null and undefined values', () => {
      interface Item {
        name: string;
        value: string | null;
      }

      const data: Item[] = [
        { name: 'Item1', value: null },
        { name: 'Item2', value: 'hello' },
      ];

      const wb = Workbook.create();
      const sheet = wb.addSheetFromData({
        name: 'Items',
        data,
      });

      expect(sheet.cell('B2').value).toBeNull();
      expect(sheet.cell('B3').value).toBe('hello');
    });

    it('preserves data after save/load cycle', async () => {
      const wb = Workbook.create();
      wb.addSheetFromData({
        name: 'People',
        data: testData,
      });

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const sheet = loaded.sheet('People');

      // Check headers
      expect(sheet.cell('A1').value).toBe('name');
      expect(sheet.cell('B1').value).toBe('age');
      expect(sheet.cell('C1').value).toBe('city');

      // Check data
      expect(sheet.cell('A2').value).toBe('Alice');
      expect(sheet.cell('B2').value).toBe(30);
      expect(sheet.cell('C2').value).toBe('Paris');
    });

    it('throws when sheet name already exists', () => {
      const wb = Workbook.create();
      wb.addSheet('Existing');
      expect(() =>
        wb.addSheetFromData({
          name: 'Existing',
          data: testData,
        }),
      ).toThrow('Sheet already exists');
    });

    describe('RichCellValue support', () => {
      it('handles rich cell values with value property', () => {
        const wb = Workbook.create();
        const data = [
          { name: 'Item', price: { value: 100 } },
          { name: 'Other', price: { value: 200 } },
        ];

        const sheet = wb.addSheetFromData({
          name: 'RichValues',
          data,
        });

        expect(sheet.cell('B2').value).toBe(100);
        expect(sheet.cell('B3').value).toBe(200);
      });

      it('handles rich cell values with formula property', () => {
        const wb = Workbook.create();
        const data = [
          { product: 'Widget', price: 10, qty: 5, total: { formula: 'B2*C2' } },
          { product: 'Gadget', price: 20, qty: 3, total: { formula: 'B3*C3' } },
        ];

        const sheet = wb.addSheetFromData({
          name: 'Formulas',
          data,
        });

        expect(sheet.cell('D2').formula).toBe('B2*C2');
        expect(sheet.cell('D3').formula).toBe('B3*C3');
      });

      it('handles rich cell values with style property', () => {
        const wb = Workbook.create();
        const data = [
          { name: 'Important', status: { value: 'OK', style: { bold: true } } },
          { name: 'Normal', status: { value: 'Pending', style: { italic: true } } },
        ];

        const sheet = wb.addSheetFromData({
          name: 'Styled',
          data,
        });

        expect(sheet.cell('B2').value).toBe('OK');
        expect(sheet.cell('B2').style?.bold).toBe(true);
        expect(sheet.cell('B3').value).toBe('Pending');
        expect(sheet.cell('B3').style?.italic).toBe(true);
      });

      it('handles rich cell values with formula and style combined', () => {
        const wb = Workbook.create();
        const data = [{ product: 'Widget', price: 10, qty: 5, total: { formula: 'B2*C2', style: { bold: true } } }];

        const sheet = wb.addSheetFromData({
          name: 'Combined',
          data,
        });

        expect(sheet.cell('D2').formula).toBe('B2*C2');
        expect(sheet.cell('D2').style?.bold).toBe(true);
      });

      it('handles rich cell values with value, formula, and style', () => {
        const wb = Workbook.create();
        const data = [{ name: 'Test', result: { value: 50, formula: 'A2*2', style: { bold: true } } }];

        const sheet = wb.addSheetFromData({
          name: 'All',
          data,
        });

        // When both value and formula are set, formula takes precedence for calculation
        // but the value may be a cached result
        expect(sheet.cell('B2').formula).toBe('A2*2');
        expect(sheet.cell('B2').style?.bold).toBe(true);
      });

      it('does not treat regular objects as rich cell values', () => {
        const wb = Workbook.create();
        const data = [{ name: 'Item', details: { color: 'red', size: 'large' } }];

        const sheet = wb.addSheetFromData({
          name: 'Objects',
          data,
        });

        // Regular objects without value/formula/style should be converted to string
        const cellValue = sheet.cell('B2').value;
        expect(typeof cellValue).toBe('string');
      });

      it('handles Date values in rich cells', () => {
        const wb = Workbook.create();
        const testDate = new Date('2024-06-15');
        const data = [{ name: 'Event', date: { value: testDate, style: { bold: true } } }];

        const sheet = wb.addSheetFromData({
          name: 'Dates',
          data,
        });

        const cellValue = sheet.cell('B2').value;
        expect(cellValue).toBeInstanceOf(Date);
        expect((cellValue as Date).getFullYear()).toBe(2024);
        expect(sheet.cell('B2').style?.bold).toBe(true);
      });

      it('preserves rich cell values through save and load', async () => {
        const wb = Workbook.create();
        const data = [{ product: 'Widget', price: 10, qty: 5, total: { formula: 'B2*C2', style: { bold: true } } }];

        wb.addSheetFromData({
          name: 'SaveLoad',
          data,
        });

        const buffer = await wb.toBuffer();
        const loaded = await Workbook.fromBuffer(buffer);
        const sheet = loaded.sheet('SaveLoad');

        expect(sheet.cell('D2').formula).toBe('B2*C2');
        expect(sheet.cell('D2').style?.bold).toBe(true);
      });
    });
  });

  describe('toJson', () => {
    it('converts sheet data to array of objects using first row as headers', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      // Set up headers
      sheet.cell('A1').value = 'name';
      sheet.cell('B1').value = 'age';
      sheet.cell('C1').value = 'city';

      // Set up data
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 30;
      sheet.cell('C2').value = 'Paris';

      sheet.cell('A3').value = 'Bob';
      sheet.cell('B3').value = 25;
      sheet.cell('C3').value = 'London';

      const result = sheet.toJson();

      expect(result).toEqual([
        { name: 'Alice', age: 30, city: 'Paris' },
        { name: 'Bob', age: 25, city: 'London' },
      ]);
    });

    it('uses provided field names instead of first row', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      // Data without headers
      sheet.cell('A1').value = 'Alice';
      sheet.cell('B1').value = 30;
      sheet.cell('C1').value = 'Paris';

      sheet.cell('A2').value = 'Bob';
      sheet.cell('B2').value = 25;
      sheet.cell('C2').value = 'London';

      const result = sheet.toJson({ fields: ['name', 'age', 'city'] });

      expect(result).toEqual([
        { name: 'Alice', age: 30, city: 'Paris' },
        { name: 'Bob', age: 25, city: 'London' },
      ]);
    });

    it('handles startRow option with headers', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      // Title row (to be skipped)
      sheet.cell('A1').value = 'Employee Report';

      // Headers at row 2 (index 1)
      sheet.cell('A2').value = 'name';
      sheet.cell('B2').value = 'salary';

      // Data
      sheet.cell('A3').value = 'Alice';
      sheet.cell('B3').value = 50000;

      const result = sheet.toJson({ startRow: 1 });

      expect(result).toEqual([{ name: 'Alice', salary: 50000 }]);
    });

    it('handles startRow option with custom fields', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      // Skip first two rows
      sheet.cell('A1').value = 'Title';
      sheet.cell('A2').value = 'Subtitle';

      // Data starts at row 3 (index 2)
      sheet.cell('A3').value = 'Alice';
      sheet.cell('B3').value = 30;

      const result = sheet.toJson({ fields: ['name', 'age'], startRow: 2 });

      expect(result).toEqual([{ name: 'Alice', age: 30 }]);
    });

    it('handles startCol option', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      // Data starts at column B (index 1)
      sheet.cell('A1').value = 'Row Number';
      sheet.cell('B1').value = 'name';
      sheet.cell('C1').value = 'age';

      sheet.cell('A2').value = 1;
      sheet.cell('B2').value = 'Alice';
      sheet.cell('C2').value = 30;

      const result = sheet.toJson({ startCol: 1 });

      expect(result).toEqual([{ name: 'Alice', age: 30 }]);
    });

    it('serializes dates as excel serial values', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'date';
      sheet.cell('A2').value = new Date('2024-06-15');

      const result = sheet.toJson({ dateHandling: 'excelSerial' });

      expect(typeof result[0].date).toBe('number');
      expect((result[0].date as number) > 45000).toBe(true);
    });

    it('serializes dates as ISO strings', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'date';
      sheet.cell('A2').value = new Date('2024-06-15T10:30:00Z');

      const result = sheet.toJson({ dateHandling: 'isoString' });

      expect(result[0].date).toBe('2024-06-15T10:30:00.000Z');
    });

    it('handles endRow option', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('A3').value = 'Bob';
      sheet.cell('A4').value = 'Charlie';

      const result = sheet.toJson({ endRow: 2 });

      expect(result).toEqual([{ name: 'Alice' }, { name: 'Bob' }]);
    });

    it('handles endCol option', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('B1').value = 'age';
      sheet.cell('C1').value = 'city';

      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 30;
      sheet.cell('C2').value = 'Paris';

      const result = sheet.toJson({ endCol: 1 });

      expect(result).toEqual([{ name: 'Alice', age: 30 }]);
    });

    it('stops on empty row by default', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('A3').value = 'Bob';
      // Row 4 is empty
      sheet.cell('A5').value = 'Charlie';

      const result = sheet.toJson();

      expect(result).toEqual([{ name: 'Alice' }, { name: 'Bob' }]);
    });

    it('continues past empty rows when stopOnEmptyRow is false', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('A2').value = 'Alice';
      // Row 3 is empty
      sheet.cell('A4').value = 'Bob';

      const result = sheet.toJson({ stopOnEmptyRow: false, endRow: 3 });

      expect(result).toEqual([{ name: 'Alice' }, { name: null }, { name: 'Bob' }]);
    });

    it('returns empty array for empty sheet', () => {
      const wb = Workbook.create();
      wb.addSheet('Empty');
      const sheet = wb.sheet('Empty');

      const result = sheet.toJson();

      expect(result).toEqual([]);
    });

    it('handles boolean values', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'active';
      sheet.cell('A2').value = true;
      sheet.cell('A3').value = false;

      const result = sheet.toJson();

      expect(result).toEqual([{ active: true }, { active: false }]);
    });

    it('handles date values', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      const testDate = new Date('2024-06-15');
      sheet.cell('A1').value = 'date';
      sheet.cell('A2').value = testDate;

      const result = sheet.toJson();

      expect(result[0].date).toBeInstanceOf(Date);
      expect((result[0].date as Date).getFullYear()).toBe(2024);
    });

    it('handles null/empty cells', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('B1').value = 'age';
      sheet.cell('A2').value = 'Alice';
      // B2 is empty

      const result = sheet.toJson();

      expect(result).toEqual([{ name: 'Alice', age: null }]);
    });

    it('generates column names for empty headers', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      // B1 is empty
      sheet.cell('C1').value = 'city';

      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 30;
      sheet.cell('C2').value = 'Paris';

      const result = sheet.toJson();

      expect(result).toEqual([{ name: 'Alice', column1: 30, city: 'Paris' }]);
    });

    it('handles duplicate headers by keeping the last value', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('B1').value = 'name';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 'Bob';

      const result = sheet.toJson();

      expect(result).toEqual([{ name: 'Bob' }]);
    });

    it('works with typed generic', () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      interface Person {
        name: string | null;
        age: number | null;
      }

      sheet.cell('A1').value = 'name';
      sheet.cell('B1').value = 'age';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 30;

      const result = sheet.toJson<Person>();

      expect(result[0].name).toBe('Alice');
      expect(result[0].age).toBe(30);
    });

    it('preserves data through addSheetFromData and toJson roundtrip', () => {
      const wb = Workbook.create();

      const originalData = [
        { name: 'Alice', age: 30, city: 'Paris' },
        { name: 'Bob', age: 25, city: 'London' },
      ];

      const sheet = wb.addSheetFromData({
        name: 'People',
        data: originalData,
      });

      const result = sheet.toJson();

      expect(result).toEqual(originalData);
    });

    it('preserves data through save/load cycle', async () => {
      const wb = Workbook.create();
      wb.addSheet('Data');
      const sheet = wb.sheet('Data');

      sheet.cell('A1').value = 'name';
      sheet.cell('B1').value = 'age';
      sheet.cell('A2').value = 'Alice';
      sheet.cell('B2').value = 30;

      const buffer = await wb.toBuffer();
      const loaded = await Workbook.fromBuffer(buffer);
      const loadedSheet = loaded.sheet('Data');

      const result = loadedSheet.toJson();

      expect(result).toEqual([{ name: 'Alice', age: 30 }]);
    });
  });
});
