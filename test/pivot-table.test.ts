import { describe, it, expect, beforeEach } from 'vitest';
import { Workbook } from '../src';
import { Styles } from '../src/styles';

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

    it('marks date fields for grouping', () => {
      wb.addSheet('Dates');
      const sheet = wb.sheet('Dates');
      sheet.cell('A1').value = 'Date';
      sheet.cell('B1').value = 'Value';
      sheet.cell('A2').value = new Date('2024-03-01');
      sheet.cell('B2').value = 10;

      const pivot = wb.createPivotTable({
        name: 'DatePivot',
        source: 'Dates!A1:B2',
        target: 'DatePivotSheet!A1',
      });

      expect(pivot.cache.fields[0].isDate).toBe(true);
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

    it('adds value fields using object config', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addValueField({ field: 'Sales', aggregation: 'sum', name: 'Total Sales' });

      expect(pivot.name).toBe('SalesPivot');
    });

    it('adds value fields with object config and number format', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addValueField({ field: 'Sales', aggregation: 'sum', name: 'Total', numberFormat: '$#,##0.00' });

      const xml = pivot.toXml();
      expect(xml).toContain('numFmtId=');
    });

    it('adds value fields with object config using defaults', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addValueField({ field: 'Sales' }); // Only field is required

      const xml = pivot.toXml();
      expect(xml).toContain('Sum of Sales'); // Default name
      expect(xml).toContain('subtotal="sum"'); // Default aggregation
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

    it('supports sorting and filtering fields', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .addColumnField('Product')
        .addValueField('Sales', 'sum')
        .sortField('Region', 'asc')
        .filterField('Product', { exclude: ['Gadget'] });

      const xml = pivot.toXml();
      expect(xml).toContain('sortType="ascending"');
      expect(xml).toContain('h="1"');
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

    it('throws when sorting non-existent field', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      expect(() => pivot.sortField('NonExistent', 'asc')).toThrow('Field not found');
    });

    it('throws when sorting field not assigned to pivot table', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region');

      expect(() => pivot.sortField('Product', 'asc')).toThrow('not assigned');
    });

    it('throws when sorting value field', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addValueField('Sales', 'sum');

      expect(() => pivot.sortField('Sales', 'asc')).toThrow('only supported for row or column');
    });

    it('throws when filtering non-existent field', () => {
      const pivot = wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      });

      expect(() => pivot.filterField('NonExistent', { include: ['A'] })).toThrow('Field not found');
    });

    it('throws when filtering field not assigned to pivot table', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region');

      expect(() => pivot.filterField('Product', { include: ['Widget'] })).toThrow('not assigned');
    });

    it('throws when using both include and exclude in filter', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region');

      expect(() => pivot.filterField('Region', { include: ['North'], exclude: ['South'] })).toThrow(
        'Cannot use both include and exclude',
      );
    });

    it('supports descending sort order', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .sortField('Region', 'desc');

      const xml = pivot.toXml();
      expect(xml).toContain('sortType="descending"');
    });

    it('supports include filter', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .filterField('Region', { include: ['North'] });

      const xml = pivot.toXml();
      // South and East should be hidden (h="1"), North should not have h attribute
      expect(xml).toContain('h="1"');
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

    it('preserves date fields in the pivot cache', async () => {
      wb.addSheet('Dates');
      const dateSheet = wb.sheet('Dates');
      dateSheet.cell('A1').value = 'Date';
      dateSheet.cell('B1').value = 'Value';
      dateSheet.cell('A2').value = new Date('2024-03-01');
      dateSheet.cell('B2').value = 100;
      dateSheet.cell('A3').value = new Date('2024-03-02');
      dateSheet.cell('B3').value = 200;

      const pivot = wb.createPivotTable({
        name: 'DatePivot',
        source: 'Dates!A1:B3',
        target: 'DatePivotSheet!A1',
      });

      const buffer = await wb.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      expect(pivot.cache.fields[0].isDate).toBe(true);
      expect(pivot.cache.fields[0].isNumeric).toBe(false);
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

  describe('value field number format', () => {
    it('adds value field with number format', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00');

      expect(pivot.name).toBe('SalesPivot');
    });

    it('includes numFmtId in XML for custom formats', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00');

      const xml = pivot.toXml();
      // Custom format should get ID >= 164
      expect(xml).toMatch(/numFmtId="\d+"/);
      expect(xml).toContain('applyNumberFormats="1"');
    });

    it('uses built-in format ID for standard formats', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .addValueField('Sales', 'sum', 'Total Sales', '#,##0.00');

      const xml = pivot.toXml();
      // #,##0.00 is built-in format ID 4
      expect(xml).toContain('numFmtId="4"');
    });

    it('sets applyNumberFormats to 0 when no formats specified', () => {
      const pivot = wb
        .createPivotTable({
          name: 'SalesPivot',
          source: 'Sheet1!A1:D7',
          target: 'PivotSheet!A3',
        })
        .addRowField('Region')
        .addValueField('Sales', 'sum', 'Total Sales');

      const xml = pivot.toXml();
      expect(xml).toContain('applyNumberFormats="0"');
    });

    it('supports multiple value fields with different formats', async () => {
      wb.createPivotTable({
        name: 'SalesPivot',
        source: 'Sheet1!A1:D7',
        target: 'PivotSheet!A3',
      })
        .addRowField('Region')
        .addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00')
        .addValueField('Quantity', 'average', 'Avg Qty', '0.00');

      const buffer = await wb.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });
});

describe('Styles built-in number formats', () => {
  it('returns built-in ID for General format', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('General')).toBe(0);
  });

  it('returns built-in ID for 0 format', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('0')).toBe(1);
  });

  it('returns built-in ID for 0.00 format', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('0.00')).toBe(2);
  });

  it('returns built-in ID for #,##0 format', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('#,##0')).toBe(3);
  });

  it('returns built-in ID for #,##0.00 format', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('#,##0.00')).toBe(4);
  });

  it('returns built-in ID for percentage formats', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('0%')).toBe(9);
    expect(styles.getOrCreateNumFmtId('0.00%')).toBe(10);
  });

  it('returns built-in ID for date formats', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('mm-dd-yy')).toBe(14);
    expect(styles.getOrCreateNumFmtId('d-mmm-yy')).toBe(15);
  });

  it('returns built-in ID for text format', () => {
    const styles = Styles.createDefault();
    expect(styles.getOrCreateNumFmtId('@')).toBe(49);
  });

  it('creates custom ID for non-built-in formats', () => {
    const styles = Styles.createDefault();
    const id = styles.getOrCreateNumFmtId('$#,##0.00');
    expect(id).toBeGreaterThanOrEqual(164);
  });

  it('reuses custom ID for same format', () => {
    const styles = Styles.createDefault();
    const id1 = styles.getOrCreateNumFmtId('$#,##0.00');
    const id2 = styles.getOrCreateNumFmtId('$#,##0.00');
    expect(id1).toBe(id2);
  });

  it('creates different IDs for different custom formats', () => {
    const styles = Styles.createDefault();
    const id1 = styles.getOrCreateNumFmtId('$#,##0.00');
    const id2 = styles.getOrCreateNumFmtId('€#,##0.00');
    expect(id1).not.toBe(id2);
  });
});
