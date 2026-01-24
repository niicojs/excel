import { describe, it, expect, beforeAll } from 'vitest';
import { Workbook } from '../src';
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
    it('creates a new workbook with default sheet', () => {
      const wb = Workbook.create();
      expect(wb.sheetCount).toBe(1);
      expect(wb.sheetNames).toEqual(['Sheet1']);
    });

    it('allows accessing the default sheet', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);
      expect(sheet.name).toBe('Sheet1');
    });
  });

  describe('sheet operations', () => {
    it('adds a new sheet', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet2');
      expect(wb.sheetCount).toBe(2);
      expect(wb.sheetNames).toEqual(['Sheet1', 'Sheet2']);
    });

    it('adds a sheet at a specific index', () => {
      const wb = Workbook.create();
      wb.addSheet('First', 0);
      expect(wb.sheetNames).toEqual(['First', 'Sheet1']);
    });

    it('throws when adding a duplicate sheet name', () => {
      const wb = Workbook.create();
      expect(() => wb.addSheet('Sheet1')).toThrow('Sheet already exists');
    });

    it('deletes a sheet by name', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet2');
      wb.deleteSheet('Sheet1');
      expect(wb.sheetNames).toEqual(['Sheet2']);
    });

    it('deletes a sheet by index', () => {
      const wb = Workbook.create();
      wb.addSheet('Sheet2');
      wb.deleteSheet(0);
      expect(wb.sheetNames).toEqual(['Sheet2']);
    });

    it('throws when deleting the last sheet', () => {
      const wb = Workbook.create();
      expect(() => wb.deleteSheet(0)).toThrow('Cannot delete the last sheet');
    });

    it('renames a sheet', () => {
      const wb = Workbook.create();
      wb.renameSheet('Sheet1', 'Renamed');
      expect(wb.sheetNames).toEqual(['Renamed']);
    });

    it('copies a sheet', () => {
      const wb = Workbook.create();
      const original = wb.sheet('Sheet1');
      original.cell('A1').value = 'Hello';
      original.cell('B1').value = 42;

      const copy = wb.copySheet('Sheet1', 'Copy');
      expect(copy.cell('A1').value).toBe('Hello');
      expect(copy.cell('B1').value).toBe(42);
    });
  });

  describe('cell operations', () => {
    it('reads and writes string values', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 'Hello';
      expect(sheet.cell('A1').value).toBe('Hello');
      expect(sheet.cell('A1').type).toBe('string');
    });

    it('reads and writes number values', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 42.5;
      expect(sheet.cell('A1').value).toBe(42.5);
      expect(sheet.cell('A1').type).toBe('number');
    });

    it('reads and writes boolean values', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = true;
      expect(sheet.cell('A1').value).toBe(true);
      expect(sheet.cell('A1').type).toBe('boolean');

      sheet.cell('A2').value = false;
      expect(sheet.cell('A2').value).toBe(false);
    });

    it('reads and writes date values', () => {
      const wb = Workbook.create();
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

    it('reads and writes formulas', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.cell('A1').value = 10;
      sheet.cell('A2').value = 20;
      sheet.cell('A3').formula = 'SUM(A1:A2)';

      expect(sheet.cell('A3').formula).toBe('SUM(A1:A2)');
    });

    it('handles null values', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      expect(sheet.cell('A1').value).toBeNull();
      expect(sheet.cell('A1').type).toBe('empty');
    });

    it('handles cell addressing by row/col', () => {
      const wb = Workbook.create();
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

    it('writes a 2D array starting at a cell', () => {
      const wb = Workbook.create();
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
      const sheet = wb.sheet(0);

      sheet.mergeCells('A1:C1');
      expect(sheet.mergedCells).toContain('A1:C1');
    });

    it('merges cells with two arguments', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.mergeCells('A1', 'C1');
      expect(sheet.mergedCells).toContain('A1:C1');
    });

    it('unmerges cells', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.mergeCells('A1:C1');
      sheet.unmergeCells('A1:C1');
      expect(sheet.mergedCells).not.toContain('A1:C1');
    });
  });

  describe('styles', () => {
    it('applies bold style', () => {
      const wb = Workbook.create();
      const sheet = wb.sheet(0);

      sheet.cell('A1').style = { bold: true };
      expect(sheet.cell('A1').style.bold).toBe(true);
    });

    it('applies multiple style properties', () => {
      const wb = Workbook.create();
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
      const sheet = wb.sheet(0);

      sheet.range('A1:B2').style = { bold: true };

      expect(sheet.cell('A1').style.bold).toBe(true);
      expect(sheet.cell('B1').style.bold).toBe(true);
      expect(sheet.cell('A2').style.bold).toBe(true);
      expect(sheet.cell('B2').style.bold).toBe(true);
    });
  });

  describe('save and load', () => {
    it('saves and loads a workbook', async () => {
      // Create and populate workbook
      const wb = Workbook.create();
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
  });
});
