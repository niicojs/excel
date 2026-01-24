import { describe, it, expect, beforeEach } from 'vitest';
import { Workbook } from '../src';

describe('PivotTable', () => {
  let wb: Workbook;

  beforeEach(() => {
    wb = Workbook.create();
    wb.addSheet('Sheet1');
    const sheet = wb.sheet('Sheet1');

    // Create sample sales data
    // Headers
    sheet.cell('A1').value = 'Region';
    sheet.cell('B1').value = 'Product';
    sheet.cell('C1').value = 'Sales';
    sheet.cell('D1').value = 'Quantity';

    // Data rows
    const data = [
      ['North', 'Widget', 100, 10],
      ['North', 'Gadget', 200, 20],
      ['South', 'Widget', 150, 15],
      ['South', 'Gadget', 250, 25],
      ['East', 'Widget', 120, 12],
      ['East', 'Gadget', 180, 18],
    ];

    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      sheet.cell(`A${row}`).value = data[i][0];
      sheet.cell(`B${row}`).value = data[i][1];
      sheet.cell(`C${row}`).value = data[i][2];
      sheet.cell(`D${row}`).value = data[i][3];
    }
  });

  describe('createPivotTable', () => {
    it('creates a pivot table with basic configuration', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      expect(pivot).toBeDefined();
      expect(pivot.name).toBe('SalesPivot');
      expect(pivot.targetSheet).toBe('PivotSheet');
      expect(pivot.targetCell).toBe('A3');
    });

    it('creates the target sheet if it does not exist', () => {
      wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'NewSheet!A1',
      });

      expect(wb.sheetNames).toContain('NewSheet');
    });

    it('throws on invalid source reference format', () => {
      expect(() => {
        wb.createPivotTable({
          name: 'SalesPivot',
          source: 'InvalidRef',
          target: 'PivotSheet!A3',
        });
      }).toThrow('Invalid reference format');
    });

    it('builds cache with correct field names', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      const cache = pivot.cache;
      expect(cache.fields.length).toBe(4);
      expect(cache.fields[0].name).toBe('Region');
      expect(cache.fields[1].name).toBe('Product');
      expect(cache.fields[2].name).toBe('Sales');
      expect(cache.fields[3].name).toBe('Quantity');
    });

    it('identifies string fields with shared items', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      const cache = pivot.cache;
      // Region field should have shared items
      expect(cache.fields[0].sharedItems).toContain('North');
      expect(cache.fields[0].sharedItems).toContain('South');
      expect(cache.fields[0].sharedItems).toContain('East');
      expect(cache.fields[0].isNumeric).toBe(false);
    });

    it('identifies numeric fields', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      const cache = pivot.cache;
      // Sales field should be numeric
      expect(cache.fields[2].isNumeric).toBe(true);
      expect(cache.fields[2].sharedItems.length).toBe(0);
      expect(cache.fields[2].minValue).toBe(100);
      expect(cache.fields[2].maxValue).toBe(250);
    });
  });

  describe('fluent API', () => {
    it('adds row fields', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region');

      expect(pivot.name).toBe('SalesPivot');
    });

    it('adds column fields', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addColumnField('Product');

      expect(pivot.name).toBe('SalesPivot');
    });

    it('adds value fields with aggregation', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addValueField('Sales', 'sum', 'Total Sales');

      expect(pivot.name).toBe('SalesPivot');
    });

    it('adds filter fields', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addFilterField('Product');

      expect(pivot.name).toBe('SalesPivot');
    });

    it('supports method chaining', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .addColumnField('Product')
        .addValueField('Sales', 'sum')
        .addValueField('Quantity', 'average')
        .addFilterField('Region');

      expect(pivot.name).toBe('SalesPivot');
    });

    it('throws when adding non-existent field', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      expect(() => pivot.addRowField('NonExistent')).toThrow('Field not found');
      expect(() => pivot.addColumnField('NonExistent')).toThrow('Field not found');
      expect(() => pivot.addValueField('NonExistent', 'sum')).toThrow('Field not found');
      expect(() => pivot.addFilterField('NonExistent')).toThrow('Field not found');
    });
  });

  describe('file generation', () => {
    it('generates a valid xlsx file with pivot table', async () => {
      wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      })
        .addRowField('Region')
        .addColumnField('Product')
        .addValueField('Sales', 'sum', 'Total Sales');

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);

      // Verify it can be read back
      const wb2 = await Workbook.fromBuffer(buffer);
      expect(wb2.sheetNames).toContain('Sheet1');
      expect(wb2.sheetNames).toContain('PivotSheet');
    });

    it('generates multiple pivot tables', async () => {
      wb.createPivotTable({
        name: 'SalesPivot1',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      })
        .addRowField('Region')
        .addValueField('Sales', 'sum');

      wb.createPivotTable({
        name: 'SalesPivot2',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!G3',
      })
        .addRowField('Product')
        .addValueField('Quantity', 'average');

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });

    it('generates pivot table with all aggregation types', async () => {
      wb.createPivotTable({
        name: 'AllAggregations',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      })
        .addRowField('Region')
        .addValueField('Sales', 'sum', 'Sum')
        .addValueField('Sales', 'count', 'Count')
        .addValueField('Sales', 'average', 'Avg')
        .addValueField('Sales', 'min', 'Min')
        .addValueField('Sales', 'max', 'Max');

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
    });
  });
});
