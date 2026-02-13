import { describe, it, expect } from 'vitest';
import { Workbook, PivotTable } from '../src';
import { readZip, readZipText } from '../src/utils/zip';

describe('PivotTable', () => {
  const createWorkbookWithSource = (): Workbook => {
    const wb = Workbook.create();
    const source = wb.addSheet('Data');
    wb.addSheet('Summary');

    source.cell('A1').value = 'Region';
    source.cell('B1').value = 'Product';
    source.cell('C1').value = 'Year';
    source.cell('D1').value = 'Sales';
    source.cell('E1').value = 'Quantity';

    source.cell('A2').value = 'North';
    source.cell('B2').value = 'Widget';
    source.cell('C2').value = 2024;
    source.cell('D2').value = 1000;
    source.cell('E2').value = 10;

    source.cell('A3').value = 'South';
    source.cell('B3').value = 'Widget';
    source.cell('C3').value = 2025;
    source.cell('D3').value = 1500;
    source.cell('E3').value = 15;

    source.cell('A4').value = 'North';
    source.cell('B4').value = 'Gadget';
    source.cell('C4').value = 2024;
    source.cell('D4').value = 900;
    source.cell('E4').value = 9;

    return wb;
  };

  it('creates a pivot table and supports fluent row/value/sort API', () => {
    const wb = createWorkbookWithSource();

    const pivot = wb.createPivotTable({
      name: 'SalesPivot',
      source: 'Data!A1:E4',
      target: 'Summary!A3',
    });

    const result = pivot
      .addRowField('Region')
      .addRowField('Product')
      .addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00')
      .sortField('Region', 'asc');

    expect(pivot).toBeInstanceOf(PivotTable);
    expect(result).toBe(pivot);
    expect(wb.pivotTables.length).toBe(1);
    expect(wb.pivotTables[0]?.name).toBe('SalesPivot');
  });

  it('supports object syntax for value fields', () => {
    const wb = createWorkbookWithSource();
    const pivot = wb.createPivotTable({
      name: 'SalesPivot',
      source: 'Data!A1:E4',
      target: 'Summary!A3',
    });

    pivot.addValueField({
      field: 'Quantity',
      aggregation: 'average',
      name: 'Avg Qty',
      numberFormat: '0.00',
    });

    expect(wb.pivotTables).toHaveLength(1);
  });

  it('throws for unknown fields in row/value/sort configuration', () => {
    const wb = createWorkbookWithSource();
    const pivot = wb.createPivotTable({
      name: 'SalesPivot',
      source: 'Data!A1:E4',
      target: 'Summary!A3',
    });

    expect(() => pivot.addRowField('Unknown')).toThrow('Pivot field not found');
    expect(() => pivot.addValueField('Missing', 'sum')).toThrow('Pivot field not found');

    pivot.addFilterField('Region');
    expect(() => pivot.sortField('Region', 'desc')).toThrow('only row or column fields can be sorted');
  });

  it('handles sparse source rows and special characters when generating cache XML', async () => {
    const wb = Workbook.create();
    const source = wb.addSheet('Data');
    wb.addSheet('Summary');

    source.cell('A1').value = 'Region & Zone';
    source.cell('B1').value = 'Sales';

    source.cell('A2').value = 'North & East';
    source.cell('B2').value = 100;

    // Sparse row: A3 missing, B3 present
    source.cell('B3').value = 50;

    source.cell('A4').value = 'West <Core>';
    source.cell('B4').value = 75;

    wb.createPivotTable({
      name: 'SparsePivot',
      source: 'Data!A1:B4',
      target: 'Summary!B2',
    })
      .addRowField('Region & Zone')
      .addValueField('Sales', 'sum', 'Total Sales');

    const buffer = await wb.toBuffer();
    const files = await readZip(buffer);

    const cacheXml = readZipText(files, 'xl/pivotCache/pivotCacheDefinition1.xml');
    expect(cacheXml).toBeTruthy();
    expect(cacheXml).toContain('cacheField');
    expect(cacheXml).toContain('Region &amp; Zone');
    expect(cacheXml).toContain('North &amp; East');
    expect(cacheXml).toContain('West &lt;Core&gt;');
  });

  it('writes pivot parts, relationships, and content-types overrides', async () => {
    const wb = createWorkbookWithSource();

    wb.createPivotTable({
      name: 'SalesPivot',
      source: 'Data!A1:E4',
      target: 'Summary!A3',
    })
      .addRowField('Region')
      .addColumnField('Year')
      .addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00')
      .sortField('Region', 'asc');

    const buffer = await wb.toBuffer();
    const files = await readZip(buffer);

    const workbookXml = readZipText(files, 'xl/workbook.xml');
    expect(workbookXml).toContain('<pivotCaches');

    const workbookRels = readZipText(files, 'xl/_rels/workbook.xml.rels');
    expect(workbookRels).toContain('pivotCacheDefinition');

    const contentTypes = readZipText(files, '[Content_Types].xml');
    expect(contentTypes).toContain('pivotCacheDefinition+xml');
    expect(contentTypes).toContain('pivotTable+xml');

    const pivotCacheXml = readZipText(files, 'xl/pivotCache/pivotCacheDefinition1.xml');
    expect(pivotCacheXml).toContain('worksheetSource');
    expect(pivotCacheXml).toContain('sheet="Data"');
    expect(pivotCacheXml).toContain('r:id="rId1"');

    const pivotCacheRecordsXml = readZipText(files, 'xl/pivotCache/pivotCacheRecords1.xml');
    expect(pivotCacheRecordsXml).toContain('pivotCacheRecords');
    expect(pivotCacheRecordsXml).toContain('count="3"');

    const pivotCacheRelsXml = readZipText(files, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels');
    expect(pivotCacheRelsXml).toContain('pivotCacheRecords');

    const pivotTableXml = readZipText(files, 'xl/pivotTables/pivotTable1.xml');
    expect(pivotTableXml).toContain('pivotTableDefinition');
    expect(pivotTableXml).toContain('dataField');
    expect(pivotTableXml).toContain('subtotal="sum"');
    expect(pivotTableXml).toContain('sortType="ascending"');
    expect(pivotTableXml).toContain('rowItems');

    const sheetRels = readZipText(files, 'xl/worksheets/_rels/sheet2.xml.rels');
    expect(sheetRels).toContain('pivotTable');

    const summaryXml = readZipText(files, 'xl/worksheets/sheet2.xml');
    expect(summaryXml).toContain('pivotTableParts');
  });
});
