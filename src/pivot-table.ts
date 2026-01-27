import type { AggregationType, PivotFieldAxis, PivotValueConfig } from './types';
import type { Styles } from './styles';
import { PivotCache } from './pivot-cache';
import { createElement, stringifyXml, XmlNode } from './utils/xml';

/**
 * Internal structure for tracking field assignments
 */
interface FieldAssignment {
  fieldName: string;
  fieldIndex: number;
  axis: PivotFieldAxis;
  aggregation?: AggregationType;
  displayName?: string;
  numFmtId?: number;
}

/**
 * Represents an Excel pivot table with a fluent API for configuration.
 */
export class PivotTable {
  private _name: string;
  private _cache: PivotCache;
  private _targetSheet: string;
  private _targetCell: string;
  private _targetRow: number;
  private _targetCol: number;

  private _rowFields: FieldAssignment[] = [];
  private _columnFields: FieldAssignment[] = [];
  private _valueFields: FieldAssignment[] = [];
  private _filterFields: FieldAssignment[] = [];

  private _pivotTableIndex: number;
  private _styles: Styles | null = null;

  constructor(
    name: string,
    cache: PivotCache,
    targetSheet: string,
    targetCell: string,
    targetRow: number,
    targetCol: number,
    pivotTableIndex: number,
  ) {
    this._name = name;
    this._cache = cache;
    this._targetSheet = targetSheet;
    this._targetCell = targetCell;
    this._targetRow = targetRow;
    this._targetCol = targetCol;
    this._pivotTableIndex = pivotTableIndex;
  }

  /**
   * Get the pivot table name
   */
  get name(): string {
    return this._name;
  }

  /**
   * Get the target sheet name
   */
  get targetSheet(): string {
    return this._targetSheet;
  }

  /**
   * Get the target cell address
   */
  get targetCell(): string {
    return this._targetCell;
  }

  /**
   * Get the pivot cache
   */
  get cache(): PivotCache {
    return this._cache;
  }

  /**
   * Get the pivot table index (for file naming)
   */
  get index(): number {
    return this._pivotTableIndex;
  }

  /**
   * Set the styles reference for number format resolution
   * @internal
   */
  setStyles(styles: Styles): this {
    this._styles = styles;
    return this;
  }

  /**
   * Add a field to the row area
   * @param fieldName - Name of the source field (column header)
   */
  addRowField(fieldName: string): this {
    const fieldIndex = this._cache.getFieldIndex(fieldName);
    if (fieldIndex < 0) {
      throw new Error(`Field not found in source data: ${fieldName}`);
    }

    this._rowFields.push({
      fieldName,
      fieldIndex,
      axis: 'row',
    });

    return this;
  }

  /**
   * Add a field to the column area
   * @param fieldName - Name of the source field (column header)
   */
  addColumnField(fieldName: string): this {
    const fieldIndex = this._cache.getFieldIndex(fieldName);
    if (fieldIndex < 0) {
      throw new Error(`Field not found in source data: ${fieldName}`);
    }

    this._columnFields.push({
      fieldName,
      fieldIndex,
      axis: 'column',
    });

    return this;
  }

  /**
   * Add a field to the values area with aggregation.
   *
   * Supports two call signatures:
   * - Positional: `addValueField(fieldName, aggregation?, displayName?, numberFormat?)`
   * - Object: `addValueField({ field, aggregation?, name?, numberFormat? })`
   *
   * @example
   * // Positional arguments
   * pivot.addValueField('Sales', 'sum', 'Total Sales', '$#,##0.00');
   *
   * // Object form
   * pivot.addValueField({ field: 'Sales', aggregation: 'sum', name: 'Total Sales', numberFormat: '$#,##0.00' });
   */
  addValueField(config: PivotValueConfig): this;
  addValueField(
    fieldName: string,
    aggregation?: AggregationType,
    displayName?: string,
    numberFormat?: string,
  ): this;
  addValueField(
    fieldNameOrConfig: string | PivotValueConfig,
    aggregation: AggregationType = 'sum',
    displayName?: string,
    numberFormat?: string,
  ): this {
    // Normalize arguments to a common form
    let fieldName: string;
    let agg: AggregationType;
    let name: string | undefined;
    let format: string | undefined;

    if (typeof fieldNameOrConfig === 'object') {
      fieldName = fieldNameOrConfig.field;
      agg = fieldNameOrConfig.aggregation ?? 'sum';
      name = fieldNameOrConfig.name;
      format = fieldNameOrConfig.numberFormat;
    } else {
      fieldName = fieldNameOrConfig;
      agg = aggregation;
      name = displayName;
      format = numberFormat;
    }

    const fieldIndex = this._cache.getFieldIndex(fieldName);
    if (fieldIndex < 0) {
      throw new Error(`Field not found in source data: ${fieldName}`);
    }

    const defaultName = `${agg.charAt(0).toUpperCase() + agg.slice(1)} of ${fieldName}`;

    // Resolve numFmtId immediately if format is provided and styles are available
    let numFmtId: number | undefined;
    if (format && this._styles) {
      numFmtId = this._styles.getOrCreateNumFmtId(format);
    }

    this._valueFields.push({
      fieldName,
      fieldIndex,
      axis: 'value',
      aggregation: agg,
      displayName: name || defaultName,
      numFmtId,
    });

    return this;
  }

  /**
   * Add a field to the filter (page) area
   * @param fieldName - Name of the source field (column header)
   */
  addFilterField(fieldName: string): this {
    const fieldIndex = this._cache.getFieldIndex(fieldName);
    if (fieldIndex < 0) {
      throw new Error(`Field not found in source data: ${fieldName}`);
    }

    this._filterFields.push({
      fieldName,
      fieldIndex,
      axis: 'filter',
    });

    return this;
  }

  /**
   * Generate the pivotTableDefinition XML
   */
  toXml(): string {
    const children: XmlNode[] = [];

    // Calculate location (estimate based on fields)
    const locationRef = this._calculateLocationRef();

    // Calculate first data row/col offsets (1-based, relative to pivot table)
    // firstHeaderRow: row offset of column headers (usually 1)
    // firstDataRow: row offset where data starts (after filters and column headers)
    // firstDataCol: column offset where data starts (after row labels)
    const filterRowCount = this._filterFields.length > 0 ? this._filterFields.length + 1 : 0;
    const headerRows = this._columnFields.length > 0 ? 1 : 0;
    const firstDataRow = filterRowCount + headerRows + 1;
    const firstDataCol = this._rowFields.length > 0 ? this._rowFields.length : 1;

    const locationNode = createElement(
      'location',
      {
        ref: locationRef,
        firstHeaderRow: String(filterRowCount + 1),
        firstDataRow: String(firstDataRow),
        firstDataCol: String(firstDataCol),
      },
      [],
    );
    children.push(locationNode);

    // Build pivotFields (one per source field)
    const pivotFieldNodes: XmlNode[] = [];
    for (const cacheField of this._cache.fields) {
      const fieldNode = this._buildPivotFieldNode(cacheField.index);
      pivotFieldNodes.push(fieldNode);
    }
    children.push(createElement('pivotFields', { count: String(pivotFieldNodes.length) }, pivotFieldNodes));

    // Row fields
    if (this._rowFields.length > 0) {
      const rowFieldNodes = this._rowFields.map((f) => createElement('field', { x: String(f.fieldIndex) }, []));
      children.push(createElement('rowFields', { count: String(rowFieldNodes.length) }, rowFieldNodes));

      // Row items
      const rowItemNodes = this._buildRowItems();
      children.push(createElement('rowItems', { count: String(rowItemNodes.length) }, rowItemNodes));
    }

    // Column fields
    if (this._columnFields.length > 0) {
      const colFieldNodes = this._columnFields.map((f) => createElement('field', { x: String(f.fieldIndex) }, []));
      // If we have multiple value fields, add -2 to indicate where "Values" header goes
      if (this._valueFields.length > 1) {
        colFieldNodes.push(createElement('field', { x: '-2' }, []));
      }
      children.push(createElement('colFields', { count: String(colFieldNodes.length) }, colFieldNodes));

      // Column items - need to account for multiple value fields
      const colItemNodes = this._buildColItems();
      children.push(createElement('colItems', { count: String(colItemNodes.length) }, colItemNodes));
    } else if (this._valueFields.length > 1) {
      // If no column fields but we have multiple values, need colFields with -2 (data field indicator)
      children.push(createElement('colFields', { count: '1' }, [createElement('field', { x: '-2' }, [])]));

      // Column items for each value field
      const colItemNodes: XmlNode[] = [];
      for (let i = 0; i < this._valueFields.length; i++) {
        colItemNodes.push(createElement('i', {}, [createElement('x', i === 0 ? {} : { v: String(i) }, [])]));
      }
      children.push(createElement('colItems', { count: String(colItemNodes.length) }, colItemNodes));
    } else if (this._valueFields.length === 1) {
      // Single value field - just add a single column item
      children.push(createElement('colItems', { count: '1' }, [createElement('i', {}, [])]));
    }

    // Page (filter) fields
    if (this._filterFields.length > 0) {
      const pageFieldNodes = this._filterFields.map((f) =>
        createElement('pageField', { fld: String(f.fieldIndex), hier: '-1' }, []),
      );
      children.push(createElement('pageFields', { count: String(pageFieldNodes.length) }, pageFieldNodes));
    }

    // Data fields (values)
    if (this._valueFields.length > 0) {
      const dataFieldNodes = this._valueFields.map((f) => {
        const attrs: Record<string, string> = {
          name: f.displayName || f.fieldName,
          fld: String(f.fieldIndex),
          baseField: '0',
          baseItem: '0',
          subtotal: f.aggregation || 'sum',
        };

        // Add numFmtId if it was resolved during addValueField
        if (f.numFmtId !== undefined) {
          attrs.numFmtId = String(f.numFmtId);
        }

        return createElement('dataField', attrs, []);
      });
      children.push(createElement('dataFields', { count: String(dataFieldNodes.length) }, dataFieldNodes));
    }

    // Check if any value field has a number format
    const hasNumberFormats = this._valueFields.some((f) => f.numFmtId !== undefined);

    // Pivot table style
    children.push(
      createElement(
        'pivotTableStyleInfo',
        {
          name: 'PivotStyleMedium9',
          showRowHeaders: '1',
          showColHeaders: '1',
          showRowStripes: '0',
          showColStripes: '0',
          showLastColumn: '1',
        },
        [],
      ),
    );

    const pivotTableNode = createElement(
      'pivotTableDefinition',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        name: this._name,
        cacheId: String(this._cache.cacheId),
        applyNumberFormats: hasNumberFormats ? '1' : '0',
        applyBorderFormats: '0',
        applyFontFormats: '0',
        applyPatternFormats: '0',
        applyAlignmentFormats: '0',
        applyWidthHeightFormats: '1',
        dataCaption: 'Values',
        updatedVersion: '8',
        minRefreshableVersion: '3',
        useAutoFormatting: '1',
        rowGrandTotals: '1',
        colGrandTotals: '1',
        itemPrintTitles: '1',
        createdVersion: '8',
        indent: '0',
        outline: '1',
        outlineData: '1',
        multipleFieldFilters: '0',
      },
      children,
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([pivotTableNode])}`;
  }

  /**
   * Build a pivotField node for a given field index
   */
  private _buildPivotFieldNode(fieldIndex: number): XmlNode {
    const attrs: Record<string, string> = {};
    const children: XmlNode[] = [];

    // Check if this field is assigned to an axis
    const rowField = this._rowFields.find((f) => f.fieldIndex === fieldIndex);
    const colField = this._columnFields.find((f) => f.fieldIndex === fieldIndex);
    const filterField = this._filterFields.find((f) => f.fieldIndex === fieldIndex);
    const valueField = this._valueFields.find((f) => f.fieldIndex === fieldIndex);

    if (rowField) {
      attrs.axis = 'axisRow';
      attrs.showAll = '0';
      // Add items for shared values
      const cacheField = this._cache.fields[fieldIndex];
      if (cacheField && cacheField.sharedItems.length > 0) {
        const itemNodes: XmlNode[] = [];
        for (let i = 0; i < cacheField.sharedItems.length; i++) {
          itemNodes.push(createElement('item', { x: String(i) }, []));
        }
        // Add default subtotal item
        itemNodes.push(createElement('item', { t: 'default' }, []));
        children.push(createElement('items', { count: String(itemNodes.length) }, itemNodes));
      }
    } else if (colField) {
      attrs.axis = 'axisCol';
      attrs.showAll = '0';
      const cacheField = this._cache.fields[fieldIndex];
      if (cacheField && cacheField.sharedItems.length > 0) {
        const itemNodes: XmlNode[] = [];
        for (let i = 0; i < cacheField.sharedItems.length; i++) {
          itemNodes.push(createElement('item', { x: String(i) }, []));
        }
        itemNodes.push(createElement('item', { t: 'default' }, []));
        children.push(createElement('items', { count: String(itemNodes.length) }, itemNodes));
      }
    } else if (filterField) {
      attrs.axis = 'axisPage';
      attrs.showAll = '0';
      const cacheField = this._cache.fields[fieldIndex];
      if (cacheField && cacheField.sharedItems.length > 0) {
        const itemNodes: XmlNode[] = [];
        for (let i = 0; i < cacheField.sharedItems.length; i++) {
          itemNodes.push(createElement('item', { x: String(i) }, []));
        }
        itemNodes.push(createElement('item', { t: 'default' }, []));
        children.push(createElement('items', { count: String(itemNodes.length) }, itemNodes));
      }
    } else if (valueField) {
      attrs.dataField = '1';
      attrs.showAll = '0';
    } else {
      attrs.showAll = '0';
    }

    return createElement('pivotField', attrs, children);
  }

  /**
   * Build row items based on unique values in row fields
   */
  private _buildRowItems(): XmlNode[] {
    const items: XmlNode[] = [];

    if (this._rowFields.length === 0) return items;

    // Get unique values from first row field
    const firstRowField = this._rowFields[0];
    const cacheField = this._cache.fields[firstRowField.fieldIndex];

    if (cacheField && cacheField.sharedItems.length > 0) {
      for (let i = 0; i < cacheField.sharedItems.length; i++) {
        items.push(createElement('i', {}, [createElement('x', i === 0 ? {} : { v: String(i) }, [])]));
      }
    }

    // Add grand total row
    items.push(createElement('i', { t: 'grand' }, [createElement('x', {}, [])]));

    return items;
  }

  /**
   * Build column items based on unique values in column fields
   */
  private _buildColItems(): XmlNode[] {
    const items: XmlNode[] = [];

    if (this._columnFields.length === 0) return items;

    // Get unique values from first column field
    const firstColField = this._columnFields[0];
    const cacheField = this._cache.fields[firstColField.fieldIndex];

    if (cacheField && cacheField.sharedItems.length > 0) {
      if (this._valueFields.length > 1) {
        // Multiple value fields - need nested items for each column value + value field combination
        for (let colIdx = 0; colIdx < cacheField.sharedItems.length; colIdx++) {
          for (let valIdx = 0; valIdx < this._valueFields.length; valIdx++) {
            const xNodes: XmlNode[] = [
              createElement('x', colIdx === 0 ? {} : { v: String(colIdx) }, []),
              createElement('x', valIdx === 0 ? {} : { v: String(valIdx) }, []),
            ];
            items.push(createElement('i', {}, xNodes));
          }
        }
      } else {
        // Single value field - simple column items
        for (let i = 0; i < cacheField.sharedItems.length; i++) {
          items.push(createElement('i', {}, [createElement('x', i === 0 ? {} : { v: String(i) }, [])]));
        }
      }
    }

    // Add grand total column(s)
    if (this._valueFields.length > 1) {
      // Grand total for each value field
      for (let valIdx = 0; valIdx < this._valueFields.length; valIdx++) {
        const xNodes: XmlNode[] = [
          createElement('x', {}, []),
          createElement('x', valIdx === 0 ? {} : { v: String(valIdx) }, []),
        ];
        items.push(createElement('i', { t: 'grand' }, xNodes));
      }
    } else {
      items.push(createElement('i', { t: 'grand' }, [createElement('x', {}, [])]));
    }

    return items;
  }

  /**
   * Calculate the location reference for the pivot table output
   */
  private _calculateLocationRef(): string {
    // Estimate output size based on fields
    const numRows = this._estimateRowCount();
    const numCols = this._estimateColCount();

    const startRow = this._targetRow;
    const startCol = this._targetCol;
    const endRow = startRow + numRows - 1;
    const endCol = startCol + numCols - 1;

    return `${this._colToLetter(startCol)}${startRow}:${this._colToLetter(endCol)}${endRow}`;
  }

  /**
   * Estimate number of rows in pivot table output
   */
  private _estimateRowCount(): number {
    let count = 1; // Header row

    // Add filter area rows
    count += this._filterFields.length;

    // Add row labels (unique values in row fields)
    if (this._rowFields.length > 0) {
      const firstRowField = this._rowFields[0];
      const cacheField = this._cache.fields[firstRowField.fieldIndex];
      count += (cacheField?.sharedItems.length || 1) + 1; // +1 for grand total
    } else {
      count += 1; // At least one data row
    }

    return Math.max(count, 3);
  }

  /**
   * Estimate number of columns in pivot table output
   */
  private _estimateColCount(): number {
    let count = 0;

    // Row label columns
    count += Math.max(this._rowFields.length, 1);

    // Column labels (unique values in column fields)
    if (this._columnFields.length > 0) {
      const firstColField = this._columnFields[0];
      const cacheField = this._cache.fields[firstColField.fieldIndex];
      count += (cacheField?.sharedItems.length || 1) + 1; // +1 for grand total
    } else {
      // Value columns
      count += Math.max(this._valueFields.length, 1);
    }

    return Math.max(count, 2);
  }

  /**
   * Convert 0-based column index to letter (A, B, ..., Z, AA, etc.)
   */
  private _colToLetter(col: number): string {
    let result = '';
    let n = col;
    while (n >= 0) {
      result = String.fromCharCode((n % 26) + 65) + result;
      n = Math.floor(n / 26) - 1;
    }
    return result;
  }
}
