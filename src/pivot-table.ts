import { PivotCache } from './pivot-cache';
import type { Styles } from './styles';
import type { AggregationType, PivotFieldAxis, PivotFieldFilter, PivotSortOrder, PivotValueConfig } from './types';
import { createElement, stringifyXml, XmlNode } from './utils/xml';

interface FieldAssignment {
  fieldName: string;
  fieldIndex: number;
  axis: PivotFieldAxis;
  aggregation?: AggregationType;
  displayName?: string;
  numFmtId?: number;
  sortOrder?: PivotSortOrder;
  filter?: PivotFieldFilter;
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
  private _pivotTableIndex: number;
  private _cacheFileIndex: number;
  private _styles: Styles | null = null;

  private _rowFields: FieldAssignment[] = [];
  private _columnFields: FieldAssignment[] = [];
  private _valueFields: FieldAssignment[] = [];
  private _filterFields: FieldAssignment[] = [];
  private _fieldAssignments: Map<number, FieldAssignment> = new Map();

  constructor(
    name: string,
    cache: PivotCache,
    targetSheet: string,
    targetCell: string,
    targetRow: number,
    targetCol: number,
    pivotTableIndex: number,
    cacheFileIndex: number,
  ) {
    this._name = name;
    this._cache = cache;
    this._targetSheet = targetSheet;
    this._targetCell = targetCell;
    this._targetRow = targetRow;
    this._targetCol = targetCol;
    this._pivotTableIndex = pivotTableIndex;
    this._cacheFileIndex = cacheFileIndex;
  }

  get name(): string {
    return this._name;
  }

  get targetSheet(): string {
    return this._targetSheet;
  }

  get targetCell(): string {
    return this._targetCell;
  }

  get cache(): PivotCache {
    return this._cache;
  }

  get index(): number {
    return this._pivotTableIndex;
  }

  get cacheFileIndex(): number {
    return this._cacheFileIndex;
  }

  setStyles(styles: Styles): this {
    this._styles = styles;
    return this;
  }

  addRowField(fieldName: string): this {
    this._addAxisField(fieldName, 'row');
    return this;
  }

  addColumnField(fieldName: string): this {
    this._addAxisField(fieldName, 'column');
    return this;
  }

  addValueField(config: PivotValueConfig): this;
  addValueField(fieldName: string, aggregation?: AggregationType, displayName?: string, numberFormat?: string): this;
  addValueField(
    fieldNameOrConfig: string | PivotValueConfig,
    aggregation: AggregationType = 'sum',
    displayName?: string,
    numberFormat?: string,
  ): this {
    const fieldName = typeof fieldNameOrConfig === 'string' ? fieldNameOrConfig : fieldNameOrConfig.field;
    const resolvedAggregation = typeof fieldNameOrConfig === 'string' ? aggregation : (fieldNameOrConfig.aggregation ?? 'sum');
    const resolvedName = typeof fieldNameOrConfig === 'string' ? displayName : fieldNameOrConfig.name;
    const resolvedFormat = typeof fieldNameOrConfig === 'string' ? numberFormat : fieldNameOrConfig.numberFormat;

    const fieldIndex = this._getFieldIndex(fieldName);
    const assignment: FieldAssignment = {
      fieldName,
      fieldIndex,
      axis: 'value',
      aggregation: resolvedAggregation,
      displayName: resolvedName ?? `${this._capitalize(resolvedAggregation)} of ${fieldName}`,
      numFmtId: resolvedFormat && this._styles ? this._styles.getOrCreateNumFmtId(resolvedFormat) : undefined,
    };

    this._valueFields.push(assignment);
    this._fieldAssignments.set(fieldIndex, assignment);
    return this;
  }

  addFilterField(fieldName: string): this {
    this._addAxisField(fieldName, 'filter');
    return this;
  }

  sortField(fieldName: string, order: PivotSortOrder): this {
    const fieldIndex = this._getFieldIndex(fieldName);
    const assignment = this._fieldAssignments.get(fieldIndex);
    if (!assignment) {
      throw new Error(`Field is not assigned to pivot table: ${fieldName}`);
    }
    if (assignment.axis !== 'row' && assignment.axis !== 'column') {
      throw new Error(`Sort is only supported for row or column fields: ${fieldName}`);
    }

    assignment.sortOrder = order;
    return this;
  }

  filterField(fieldName: string, filter: PivotFieldFilter): this {
    const fieldIndex = this._getFieldIndex(fieldName);
    const assignment = this._fieldAssignments.get(fieldIndex);
    if (!assignment) {
      throw new Error(`Field is not assigned to pivot table: ${fieldName}`);
    }
    if (filter.include && filter.exclude) {
      throw new Error('Cannot use both include and exclude in the same filter');
    }

    assignment.filter = filter;
    return this;
  }

  /**
   * Generate the pivotTableDefinition XML.
   */
  toXml(): string {
    const children: XmlNode[] = [this._buildLocationNode(), this._buildPivotFieldsNode()];

    if (this._rowFields.length > 0) {
      children.push(this._buildAxisFieldsNode('rowFields', this._rowFields));
      children.push(this._buildAxisItemsNode('rowItems', this._rowFields));
    }

    if (this._columnFields.length > 0 || this._valueFields.length > 1) {
      const colFields = [...this._columnFields];
      const fieldNodes = colFields.map((field) => createElement('field', { x: String(field.fieldIndex) }, []));
      if (this._valueFields.length > 1) {
        fieldNodes.push(createElement('field', { x: '-2' }, []));
      }
      children.push(createElement('colFields', { count: String(fieldNodes.length) }, fieldNodes));
      children.push(this._buildAxisItemsNode('colItems', colFields, this._valueFields.length > 1));
    } else {
      children.push(createElement('colItems', { count: '1' }, [createElement('i', {}, [])]));
    }

    if (this._filterFields.length > 0) {
      children.push(
        createElement(
          'pageFields',
          { count: String(this._filterFields.length) },
          this._filterFields.map((field) => createElement('pageField', { fld: String(field.fieldIndex), hier: '-1' }, [])),
        ),
      );
    }

    if (this._valueFields.length > 0) {
      children.push(this._buildDataFieldsNode());
    }

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
        dataOnRows: '0',
        applyNumberFormats: this._valueFields.some((field) => field.numFmtId !== undefined) ? '1' : '0',
        applyBorderFormats: '0',
        applyFontFormats: '0',
        applyPatternFormats: '0',
        applyAlignmentFormats: '0',
        applyWidthHeightFormats: '1',
        dataCaption: 'Values',
        grandTotalCaption: 'Grand Total',
        updatedVersion: '8',
        minRefreshableVersion: '3',
        useAutoFormatting: '1',
        itemPrintTitles: '1',
        createdVersion: '8',
        indent: '0',
        compact: '1',
        compactData: '1',
        outline: '1',
        outlineData: '1',
        multipleFieldFilters: '0',
      },
      children,
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([pivotTableNode])}`;
  }

  private _addAxisField(fieldName: string, axis: 'row' | 'column' | 'filter'): void {
    const fieldIndex = this._getFieldIndex(fieldName);
    const assignment: FieldAssignment = { fieldName, fieldIndex, axis };
    if (axis === 'row') {
      this._rowFields.push(assignment);
    } else if (axis === 'column') {
      this._columnFields.push(assignment);
    } else {
      this._filterFields.push(assignment);
    }
    this._fieldAssignments.set(fieldIndex, assignment);
  }

  private _getFieldIndex(fieldName: string): number {
    const fieldIndex = this._cache.getFieldIndex(fieldName);
    if (fieldIndex < 0) {
      throw new Error(`Field not found in source data: ${fieldName}`);
    }
    return fieldIndex;
  }

  private _buildLocationNode(): XmlNode {
    const filterRows = this._filterFields.length > 0 ? this._filterFields.length + 1 : 0;
    const headerRows = this._columnFields.length > 0 || this._valueFields.length > 1 ? 1 : 0;
    const firstDataRow = filterRows + headerRows + 1;
    const firstDataCol = Math.max(this._rowFields.length, 1);

    return createElement(
      'location',
      {
        ref: this._calculateLocationRef(),
        firstHeaderRow: String(filterRows + 1),
        firstDataRow: String(firstDataRow),
        firstDataCol: String(firstDataCol),
      },
      [],
    );
  }

  private _buildPivotFieldsNode(): XmlNode {
    const fieldNodes = this._cache.fields.map((field) => this._buildPivotFieldNode(field.index));
    return createElement('pivotFields', { count: String(fieldNodes.length) }, fieldNodes);
  }

  private _buildPivotFieldNode(fieldIndex: number): XmlNode {
    const assignment = this._fieldAssignments.get(fieldIndex);
    const attrs: Record<string, string> = { showAll: '0' };
    const children: XmlNode[] = [];

    if (assignment?.axis === 'row') {
      attrs.axis = 'axisRow';
    } else if (assignment?.axis === 'column') {
      attrs.axis = 'axisCol';
    } else if (assignment?.axis === 'filter') {
      attrs.axis = 'axisPage';
    } else if (assignment?.axis === 'value') {
      attrs.dataField = '1';
    }

    if (assignment?.sortOrder) {
      attrs.sortType = assignment.sortOrder === 'asc' ? 'ascending' : 'descending';
    }

    const cacheField = this._cache.fields[fieldIndex];
    if (assignment && assignment.axis !== 'value' && cacheField?.sharedItems.length) {
      children.push(createElement('items', { count: String(cacheField.sharedItems.length + 1) }, this._buildItemNodes(fieldIndex)));
    }

    return createElement('pivotField', attrs, children);
  }

  private _buildItemNodes(fieldIndex: number): XmlNode[] {
    const assignment = this._fieldAssignments.get(fieldIndex);
    const cacheField = this._cache.fields[fieldIndex];
    const itemNodes = cacheField.sharedItems.map((value, index) => {
      const attrs: Record<string, string> = { x: String(index) };
      if (assignment?.filter && this._isHiddenByFilter(value, assignment.filter)) {
        attrs.h = '1';
      }
      return createElement('item', attrs, []);
    });
    itemNodes.push(createElement('item', { t: 'default' }, []));
    return itemNodes;
  }

  private _buildAxisFieldsNode(tagName: 'rowFields' | 'colFields', fields: FieldAssignment[]): XmlNode {
    return createElement(
      tagName,
      { count: String(fields.length) },
      fields.map((field) => createElement('field', { x: String(field.fieldIndex) }, [])),
    );
  }

  private _buildAxisItemsNode(tagName: 'rowItems' | 'colItems', fields: FieldAssignment[], includeValues = false): XmlNode {
    const itemCount = Math.max(this._getLargestSharedItemCount(fields), 1);
    const items: XmlNode[] = [];

    for (let i = 0; i < itemCount; i++) {
      const xNodes = fields.map(() => createElement('x', i === 0 ? {} : { v: String(i) }, []));
      if (includeValues) {
        for (let valueIndex = 0; valueIndex < this._valueFields.length; valueIndex++) {
          items.push(createElement('i', {}, [...xNodes, createElement('x', valueIndex === 0 ? {} : { v: String(valueIndex) }, [])]));
        }
      } else {
        items.push(createElement('i', {}, xNodes));
      }
    }

    const grandTotalNodes = fields.map(() => createElement('x', {}, []));
    if (includeValues) {
      for (let valueIndex = 0; valueIndex < this._valueFields.length; valueIndex++) {
        items.push(
          createElement('i', { t: 'grand' }, [
            ...grandTotalNodes,
            createElement('x', valueIndex === 0 ? {} : { v: String(valueIndex) }, []),
          ]),
        );
      }
    } else {
      items.push(createElement('i', { t: 'grand' }, grandTotalNodes));
    }

    return createElement(tagName, { count: String(items.length) }, items);
  }

  private _buildDataFieldsNode(): XmlNode {
    const dataFields = this._valueFields.map((field) => {
      const attrs: Record<string, string> = {
        name: field.displayName ?? field.fieldName,
        fld: String(field.fieldIndex),
        subtotal: field.aggregation ?? 'sum',
      };
      if (field.numFmtId !== undefined) {
        attrs.numFmtId = String(field.numFmtId);
      }
      return createElement('dataField', attrs, []);
    });

    return createElement('dataFields', { count: String(dataFields.length) }, dataFields);
  }

  private _isHiddenByFilter(value: string, filter: PivotFieldFilter): boolean {
    if (filter.exclude) {
      return filter.exclude.includes(value);
    }
    if (filter.include) {
      return !filter.include.includes(value);
    }
    return false;
  }

  private _getLargestSharedItemCount(fields: FieldAssignment[]): number {
    return fields.reduce((largest, field) => Math.max(largest, this._cache.fields[field.fieldIndex]?.sharedItems.length ?? 0), 0);
  }

  private _calculateLocationRef(): string {
    const startRow = this._targetRow;
    const startCol = this._targetCol;
    const endRow = startRow + this._estimateRowCount() - 1;
    const endCol = startCol + this._estimateColCount() - 1;
    return `${this._colToLetter(startCol)}${startRow}:${this._colToLetter(endCol)}${endRow}`;
  }

  private _estimateRowCount(): number {
    const filterRows = this._filterFields.length > 0 ? this._filterFields.length + 1 : 0;
    const headerRows = this._columnFields.length > 0 || this._valueFields.length > 1 ? 1 : 0;
    const itemRows = Math.max(this._getLargestSharedItemCount(this._rowFields), 1) + 1;
    return Math.max(filterRows + headerRows + itemRows, 3);
  }

  private _estimateColCount(): number {
    const rowColumns = Math.max(this._rowFields.length, 1);
    const columnItems = this._columnFields.length > 0 ? Math.max(this._getLargestSharedItemCount(this._columnFields), 1) + 1 : 1;
    const valueMultiplier = this._columnFields.length > 0 ? Math.max(this._valueFields.length, 1) : Math.max(this._valueFields.length, 1);
    return Math.max(rowColumns + columnItems * valueMultiplier, 2);
  }

  private _colToLetter(col: number): string {
    let result = '';
    let current = col;
    while (current >= 0) {
      result = String.fromCharCode((current % 26) + 65) + result;
      current = Math.floor(current / 26) - 1;
    }
    return result;
  }

  private _capitalize(value: string): string {
    return `${value.charAt(0).toUpperCase()}${value.slice(1)}`;
  }
}
