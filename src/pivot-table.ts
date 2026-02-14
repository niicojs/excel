import type {
  PivotAggregationType,
  PivotFieldFilter,
  PivotSortOrder,
  PivotTableConfig,
  PivotValueConfig,
  RangeAddress,
  CellValue,
} from './types';
import type { Workbook } from './workbook';
import type { Worksheet } from './worksheet';
import { createElement, stringifyXml, XmlNode } from './utils/xml';
import { toAddress, toRange } from './utils/address';

const AGGREGATION_TO_XML: Record<PivotAggregationType, string> = {
  sum: 'sum',
  count: 'count',
  average: 'average',
  min: 'min',
  max: 'max',
};

const SORT_TO_XML: Record<PivotSortOrder, 'ascending' | 'descending'> = {
  asc: 'ascending',
  desc: 'descending',
};

interface PivotValueField {
  field: string;
  aggregation: PivotAggregationType;
  name: string;
  numberFormat?: string;
}

interface PivotFieldMeta {
  name: string;
  sourceCol: number;
}

/**
 * Represents an Excel PivotTable with a fluent configuration API.
 */
export class PivotTable {
  private _workbook: Workbook;
  private _name: string;
  private _sourceSheetName: string;
  private _sourceSheet: Worksheet;
  private _sourceRange: RangeAddress;
  private _targetSheetName: string;
  private _targetCell: { row: number; col: number };
  private _refreshOnLoad: boolean;
  private _cacheId: number;
  private _pivotId: number;
  private _cachePartIndex: number;
  private _fields: PivotFieldMeta[];

  private _rowFields: string[] = [];
  private _columnFields: string[] = [];
  private _filterFields: string[] = [];
  private _valueFields: PivotValueField[] = [];
  private _sortOrders: Map<string, PivotSortOrder> = new Map();
  private _filters: Map<string, PivotFieldFilter> = new Map();

  constructor(
    workbook: Workbook,
    config: PivotTableConfig,
    sourceSheetName: string,
    sourceSheet: Worksheet,
    sourceRange: RangeAddress,
    targetSheetName: string,
    targetCell: { row: number; col: number },
    cacheId: number,
    pivotId: number,
    cachePartIndex: number,
    fields: PivotFieldMeta[],
  ) {
    this._workbook = workbook;
    this._name = config.name;
    this._sourceSheetName = sourceSheetName;
    this._sourceSheet = sourceSheet;
    this._sourceRange = sourceRange;
    this._targetSheetName = targetSheetName;
    this._targetCell = targetCell;
    this._refreshOnLoad = config.refreshOnLoad !== false;
    this._cacheId = cacheId;
    this._pivotId = pivotId;
    this._cachePartIndex = cachePartIndex;
    this._fields = fields;
  }

  get name(): string {
    return this._name;
  }

  get sourceSheetName(): string {
    return this._sourceSheetName;
  }

  get sourceRange(): RangeAddress {
    return { start: { ...this._sourceRange.start }, end: { ...this._sourceRange.end } };
  }

  get targetSheetName(): string {
    return this._targetSheetName;
  }

  get targetCell(): { row: number; col: number } {
    return { ...this._targetCell };
  }

  get refreshOnLoad(): boolean {
    return this._refreshOnLoad;
  }

  get cacheId(): number {
    return this._cacheId;
  }

  get pivotId(): number {
    return this._pivotId;
  }

  get cachePartIndex(): number {
    return this._cachePartIndex;
  }

  addRowField(fieldName: string): this {
    this._assertFieldExists(fieldName);
    if (!this._rowFields.includes(fieldName)) {
      this._rowFields.push(fieldName);
    }
    return this;
  }

  addColumnField(fieldName: string): this {
    this._assertFieldExists(fieldName);
    if (!this._columnFields.includes(fieldName)) {
      this._columnFields.push(fieldName);
    }
    return this;
  }

  addFilterField(fieldName: string): this {
    this._assertFieldExists(fieldName);
    if (!this._filterFields.includes(fieldName)) {
      this._filterFields.push(fieldName);
    }
    return this;
  }

  addValueField(
    fieldName: string,
    aggregation?: PivotAggregationType,
    displayName?: string,
    numberFormat?: string,
  ): this;
  addValueField(config: PivotValueConfig): this;
  addValueField(
    fieldNameOrConfig: string | PivotValueConfig,
    aggregation: PivotAggregationType = 'sum',
    displayName?: string,
    numberFormat?: string,
  ): this {
    let config: PivotValueConfig;

    if (typeof fieldNameOrConfig === 'string') {
      config = {
        field: fieldNameOrConfig,
        aggregation,
        name: displayName,
        numberFormat,
      };
    } else {
      config = fieldNameOrConfig;
    }

    this._assertFieldExists(config.field);

    const resolvedAggregation = config.aggregation ?? 'sum';
    const resolvedName = config.name ?? `${this._aggregationLabel(resolvedAggregation)} of ${config.field}`;

    this._valueFields.push({
      field: config.field,
      aggregation: resolvedAggregation,
      name: resolvedName,
      numberFormat: config.numberFormat,
    });

    return this;
  }

  sortField(fieldName: string, order: PivotSortOrder): this {
    this._assertFieldExists(fieldName);
    if (!this._rowFields.includes(fieldName) && !this._columnFields.includes(fieldName)) {
      throw new Error(`Cannot sort field "${fieldName}": only row or column fields can be sorted`);
    }
    this._sortOrders.set(fieldName, order);
    return this;
  }

  filterField(fieldName: string, filter: PivotFieldFilter): this {
    this._assertFieldExists(fieldName);

    const hasInclude = 'include' in filter;
    const hasExclude = 'exclude' in filter;
    if ((hasInclude && hasExclude) || (!hasInclude && !hasExclude)) {
      throw new Error('Pivot filter must contain either include or exclude');
    }

    const values = hasInclude ? filter.include : filter.exclude;
    if (!values || values.length === 0) {
      throw new Error('Pivot filter values cannot be empty');
    }

    this._filters.set(fieldName, filter);
    return this;
  }

  toPivotCacheDefinitionXml(): string {
    const rowCount = Math.max(0, this._sourceRange.end.row - this._sourceRange.start.row);
    const cacheFieldNodes = this._fields.map((field, index) => this._buildCacheFieldNode(field, index));

    const attrs: Record<string, string> = {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'mc:Ignorable': 'xr',
      'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
      'r:id': 'rId1',
      createdVersion: '8',
      minRefreshableVersion: '3',
      refreshedVersion: '8',
      refreshOnLoad: this._refreshOnLoad ? '1' : '0',
      recordCount: String(rowCount),
    };

    const cacheSourceNode = createElement('cacheSource', { type: 'worksheet' }, [
      createElement('worksheetSource', { sheet: this._sourceSheetName, ref: toRange(this._sourceRange) }, []),
    ]);

    const cacheFieldsNode = createElement('cacheFields', { count: String(cacheFieldNodes.length) }, cacheFieldNodes);

    const extLstNode = createElement('extLst', {}, [
      createElement(
        'ext',
        {
          uri: '{725AE2AE-9491-48be-B2B4-4EB974FC3084}',
          'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
        },
        [createElement('x14:pivotCacheDefinition', {}, [])],
      ),
    ]);

    const root = createElement('pivotCacheDefinition', attrs, [cacheSourceNode, cacheFieldsNode, extLstNode]);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([root])}`;
  }

  toPivotCacheRecordsXml(): string {
    const recordNodes: XmlNode[] = [];
    for (let row = this._sourceRange.start.row + 1; row <= this._sourceRange.end.row; row++) {
      const valueNodes: XmlNode[] = [];

      for (let fieldIndex = 0; fieldIndex < this._fields.length; fieldIndex++) {
        const field = this._fields[fieldIndex];
        const cellValue = this._sourceSheet.getCellIfExists(row, field.sourceCol)?.value ?? null;
        valueNodes.push(this._toCacheRecordValueNode(cellValue, fieldIndex));
      }

      recordNodes.push(createElement('r', {}, valueNodes));
    }

    const root = createElement(
      'pivotCacheRecords',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'mc:Ignorable': 'xr',
        'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
        count: String(recordNodes.length),
      },
      recordNodes,
    );
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([root])}`;
  }

  toPivotCacheDefinitionRelsXml(): string {
    const relsRoot = createElement(
      'Relationships',
      { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
      [
        createElement(
          'Relationship',
          {
            Id: 'rId1',
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
            Target: `pivotCacheRecords${this._cachePartIndex}.xml`,
          },
          [],
        ),
      ],
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([relsRoot])}`;
  }

  toPivotTableDefinitionXml(): string {
    const effectiveValueFields = this._valueFields.length > 0 ? [this._valueFields[0]] : [];
    const sourceFieldCount = this._fields.length;
    const pivotFields: XmlNode[] = [];
    const effectiveRowFieldName = this._rowFields[0];
    const rowFieldIndexes = effectiveRowFieldName ? [this._fieldIndex(effectiveRowFieldName)] : [];
    const colFieldIndexes = this._columnFields.length > 0 ? [this._fieldIndex(this._columnFields[0])] : [];
    const valueFieldIndexes = new Set<number>(
      effectiveValueFields.map((valueField) => this._fieldIndex(valueField.field)),
    );

    for (let index = 0; index < this._fields.length; index++) {
      const field = this._fields[index];
      const attrs: Record<string, string> = { showAll: '0' };

      if (rowFieldIndexes.includes(index)) {
        attrs.axis = 'axisRow';
      } else if (colFieldIndexes.includes(index)) {
        attrs.axis = 'axisCol';
      }

      if (valueFieldIndexes.has(index)) {
        attrs.dataField = '1';
      }

      const sortOrder = this._sortOrders.get(field.name);
      if (sortOrder) {
        attrs.sortType = SORT_TO_XML[sortOrder];
      }

      const children: XmlNode[] = [];
      if (rowFieldIndexes.includes(index) || colFieldIndexes.includes(index)) {
        const distinctItems = this._collectDistinctItems(index);
        const itemNodes: XmlNode[] = distinctItems.map((_item, itemIndex) =>
          createElement('item', { x: String(itemIndex) }, []),
        );
        itemNodes.push(createElement('item', { t: 'default' }, []));
        children.push(createElement('items', { count: String(itemNodes.length) }, itemNodes));
      }

      pivotFields.push(createElement('pivotField', attrs, children));
    }

    const children: XmlNode[] = [];

    const locationRef = this._buildTargetAreaRef();
    children.push(
      createElement(
        'location',
        {
          ref: locationRef,
          firstHeaderRow: '1',
          firstDataRow: '1',
          firstDataCol: String(Math.max(1, this._rowFields.length + 1)),
        },
        [],
      ),
    );

    children.push(createElement('pivotFields', { count: String(sourceFieldCount) }, pivotFields));

    if (rowFieldIndexes.length > 0) {
      children.push(
        createElement(
          'rowFields',
          { count: String(rowFieldIndexes.length) },
          rowFieldIndexes.map((fieldIndex) => createElement('field', { x: String(fieldIndex) }, [])),
        ),
      );

      const distinctRowItems = this._collectDistinctItems(rowFieldIndexes[0]);
      const rowItemNodes: XmlNode[] = [];
      if (distinctRowItems.length > 0) {
        rowItemNodes.push(createElement('i', {}, [createElement('x', {}, [])]));
        for (let itemIndex = 1; itemIndex < distinctRowItems.length; itemIndex++) {
          rowItemNodes.push(createElement('i', {}, [createElement('x', { v: String(itemIndex) }, [])]));
        }
      }
      rowItemNodes.push(createElement('i', { t: 'grand' }, [createElement('x', {}, [])]));
      children.push(createElement('rowItems', { count: String(rowItemNodes.length) }, rowItemNodes));
    }

    if (colFieldIndexes.length > 0) {
      children.push(
        createElement(
          'colFields',
          { count: String(colFieldIndexes.length) },
          colFieldIndexes.map((fieldIndex) => createElement('field', { x: String(fieldIndex) }, [])),
        ),
      );
    }

    // Excel expects colItems even when no explicit column fields are configured.
    children.push(createElement('colItems', { count: '1' }, [createElement('i', {}, [])]));

    if (this._filterFields.length > 0) {
      children.push(
        createElement(
          'pageFields',
          { count: String(this._filterFields.length) },
          this._filterFields.map((field, index) =>
            createElement('pageField', { fld: String(this._fieldIndex(field)), hier: '-1', item: String(index) }, []),
          ),
        ),
      );
    }

    if (effectiveValueFields.length > 0) {
      children.push(
        createElement(
          'dataFields',
          { count: String(effectiveValueFields.length) },
          effectiveValueFields.map((valueField) => {
            const attrs: Record<string, string> = {
              name: valueField.name,
              fld: String(this._fieldIndex(valueField.field)),
              baseField: '0',
              baseItem: '0',
              subtotal: AGGREGATION_TO_XML[valueField.aggregation],
            };

            if (valueField.numberFormat) {
              attrs.numFmtId = String(this._workbook.styles.getOrCreateNumFmtId(valueField.numberFormat));
            }

            return createElement('dataField', attrs, []);
          }),
        ),
      );
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

    const attrs: Record<string, string> = {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      name: this._name,
      cacheId: String(this._cacheId),
      dataCaption: 'Values',
      applyNumberFormats: '1',
      applyBorderFormats: '0',
      applyFontFormats: '0',
      applyPatternFormats: '0',
      applyAlignmentFormats: '0',
      applyWidthHeightFormats: '1',
      updatedVersion: '8',
      minRefreshableVersion: '3',
      createdVersion: '8',
      useAutoFormatting: '1',
      rowGrandTotals: '1',
      colGrandTotals: '1',
      itemPrintTitles: '1',
      indent: '0',
      multipleFieldFilters: this._filters.size > 0 ? '1' : '0',
      outline: '1',
      outlineData: '1',
    };

    const root = createElement('pivotTableDefinition', attrs, children);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([root])}`;
  }

  private _buildCacheFieldNode(field: PivotFieldMeta, fieldIndex: number): XmlNode {
    const values = this._collectFieldValues(fieldIndex);
    const nonNullValues = values.filter((value): value is Exclude<CellValue, null> => value !== null);
    const isAxisField = this._isAxisField(field.name);
    const isValueField = this._isValueField(field.name);
    const numericValues = nonNullValues.filter(
      (value): value is number => typeof value === 'number' && Number.isFinite(value),
    );
    const allNonNullAreNumbers = nonNullValues.length > 0 && numericValues.length === nonNullValues.length;

    if (isValueField || (!isAxisField && allNonNullAreNumbers)) {
      const minValue = numericValues.length > 0 ? Math.min(...numericValues) : 0;
      const maxValue = numericValues.length > 0 ? Math.max(...numericValues) : 0;
      const hasInteger = numericValues.every((value) => Number.isInteger(value));

      const attrs: Record<string, string> = {
        containsSemiMixedTypes: '0',
        containsString: '0',
        containsNumber: '1',
        minValue: String(minValue),
        maxValue: String(maxValue),
      };
      if (hasInteger) {
        attrs.containsInteger = '1';
      }

      return createElement('cacheField', { name: field.name, numFmtId: '0' }, [
        createElement('sharedItems', attrs, []),
      ]);
    }

    if (!isAxisField) {
      return createElement('cacheField', { name: field.name, numFmtId: '0' }, [createElement('sharedItems', {}, [])]);
    }

    const distinct = new Set<string>();
    const sharedItems: XmlNode[] = [];

    for (const value of values) {
      if (value === null) {
        continue;
      }

      const key = this._distinctKey(value);
      if (distinct.has(key)) {
        continue;
      }
      distinct.add(key);

      if (typeof value === 'string') {
        sharedItems.push({ s: [], ':@': { '@_v': value } } as XmlNode);
        continue;
      }

      if (typeof value === 'number') {
        sharedItems.push(createElement('n', { v: String(value) }, []));
        continue;
      }

      if (typeof value === 'boolean') {
        sharedItems.push(createElement('b', { v: value ? '1' : '0' }, []));
        continue;
      }

      if (value instanceof Date) {
        sharedItems.push(createElement('d', { v: value.toISOString() }, []));
      }
    }

    return createElement('cacheField', { name: field.name, numFmtId: '0' }, [
      createElement('sharedItems', { count: String(sharedItems.length) }, sharedItems),
    ]);
  }

  private _buildTargetAreaRef(): string {
    const start = this._targetCell;
    const estimatedRows = Math.max(3, this._estimateOutputRows());
    const estimatedCols = Math.max(1, this._rowFields.length + Math.max(1, this._valueFields.length));

    const endRow = start.row + estimatedRows - 1;
    const endCol = start.col + estimatedCols - 1;

    return `${toAddress(start.row, start.col)}:${toAddress(endRow, endCol)}`;
  }

  private _estimateOutputRows(): number {
    if (this._rowFields.length === 0) {
      return 3;
    }

    const rowFieldIndex = this._fieldIndex(this._rowFields[0]);
    const values = this._collectFieldValues(rowFieldIndex)
      .filter((value): value is Exclude<CellValue, null> => value !== null)
      .map((value) => this._distinctKey(value));

    const distinct = new Set(values);
    return Math.max(3, distinct.size + 2);
  }

  private _collectFieldValues(fieldIndex: number): CellValue[] {
    const meta = this._fields[fieldIndex];
    const values: CellValue[] = [];

    for (let row = this._sourceRange.start.row + 1; row <= this._sourceRange.end.row; row++) {
      const cell = this._sourceSheet.getCellIfExists(row, meta.sourceCol);
      values.push((cell?.value ?? null) as CellValue);
    }

    return values;
  }

  private _collectDistinctItems(fieldIndex: number): Exclude<CellValue, null>[] {
    const distinct = new Set<string>();
    const result: Exclude<CellValue, null>[] = [];

    for (const value of this._collectFieldValues(fieldIndex)) {
      if (value === null) continue;
      const key = this._distinctKey(value);
      if (distinct.has(key)) continue;
      distinct.add(key);
      result.push(value);
    }

    return result;
  }

  private _toCacheRecordValueNode(value: CellValue, fieldIndex: number): XmlNode {
    if (value === null) {
      return createElement('m', {}, []);
    }

    const field = this._fields[fieldIndex];
    if (this._isAxisField(field.name)) {
      const index = this._sharedItemIndex(fieldIndex, value);
      if (index >= 0) {
        return createElement('x', { v: String(index) }, []);
      }
    }

    if (typeof value === 'string') {
      return { s: [], ':@': { '@_v': value } } as XmlNode;
    }

    if (typeof value === 'number') {
      return createElement('n', { v: String(value) }, []);
    }

    if (typeof value === 'boolean') {
      return createElement('b', { v: value ? '1' : '0' }, []);
    }

    if (value instanceof Date) {
      return createElement('d', { v: value.toISOString() }, []);
    }

    return createElement('m', {}, []);
  }

  private _assertFieldExists(fieldName: string): void {
    if (!this._fields.some((field) => field.name === fieldName)) {
      throw new Error(`Pivot field not found: ${fieldName}`);
    }
  }

  private _isValueField(fieldName: string): boolean {
    return this._valueFields.some((valueField) => valueField.field === fieldName);
  }

  private _isAxisField(fieldName: string): boolean {
    const effectiveRowField = this._rowFields[0] ?? null;
    const effectiveColumnField = this._columnFields[0] ?? null;
    return (
      fieldName === effectiveRowField || fieldName === effectiveColumnField || this._filterFields.includes(fieldName)
    );
  }

  private _sharedItemIndex(fieldIndex: number, value: Exclude<CellValue, null>): number {
    const distinct = new Map<string, number>();
    for (const item of this._collectFieldValues(fieldIndex)) {
      if (item === null) {
        continue;
      }
      const key = this._distinctKey(item);
      if (!distinct.has(key)) {
        distinct.set(key, distinct.size);
      }
    }

    const targetKey = this._distinctKey(value);
    return distinct.get(targetKey) ?? -1;
  }

  private _fieldIndex(fieldName: string): number {
    const index = this._fields.findIndex((field) => field.name === fieldName);
    if (index < 0) {
      throw new Error(`Pivot field not found: ${fieldName}`);
    }
    return index;
  }

  private _aggregationLabel(aggregation: PivotAggregationType): string {
    switch (aggregation) {
      case 'sum':
        return 'Sum';
      case 'count':
        return 'Count';
      case 'average':
        return 'Average';
      case 'min':
        return 'Min';
      case 'max':
        return 'Max';
    }
  }

  private _distinctKey(value: Exclude<CellValue, null>): string {
    if (value instanceof Date) {
      return `d:${value.toISOString()}`;
    }
    if (typeof value === 'string') {
      return `s:${value}`;
    }
    if (typeof value === 'number') {
      return `n:${value}`;
    }
    if (typeof value === 'boolean') {
      return `b:${value ? 1 : 0}`;
    }
    if (typeof value === 'object' && value && 'error' in value) {
      return `e:${value.error}`;
    }
    return 'u:';
  }
}
