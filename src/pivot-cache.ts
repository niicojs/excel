import type { PivotCacheField, CellValue } from './types';
import { createElement, stringifyXml, XmlNode } from './utils/xml';

/**
 * Manages the pivot cache (definition and records) for a pivot table.
 * The cache stores source data metadata and cached values.
 */
export class PivotCache {
  private _cacheId: number;
  private _sourceSheet: string;
  private _sourceRange: string;
  private _fields: PivotCacheField[] = [];
  private _records: CellValue[][] = [];
  private _recordCount = 0;
  private _refreshOnLoad = true; // Default to true
  private _dateGrouping = false;

  constructor(cacheId: number, sourceSheet: string, sourceRange: string) {
    this._cacheId = cacheId;
    this._sourceSheet = sourceSheet;
    this._sourceRange = sourceRange;
  }

  /**
   * Get the cache ID
   */
  get cacheId(): number {
    return this._cacheId;
  }

  /**
   * Set refreshOnLoad option
   */
  set refreshOnLoad(value: boolean) {
    this._refreshOnLoad = value;
  }

  /**
   * Get refreshOnLoad option
   */
  get refreshOnLoad(): boolean {
    return this._refreshOnLoad;
  }

  /**
   * Get the source sheet name
   */
  get sourceSheet(): string {
    return this._sourceSheet;
  }

  /**
   * Get the source range
   */
  get sourceRange(): string {
    return this._sourceRange;
  }

  /**
   * Get the full source reference (Sheet!Range)
   */
  get sourceRef(): string {
    return `${this._sourceSheet}!${this._sourceRange}`;
  }

  /**
   * Get the fields in this cache
   */
  get fields(): PivotCacheField[] {
    return this._fields;
  }

  /**
   * Get the number of data records
   */
  get recordCount(): number {
    return this._recordCount;
  }

  /**
   * Build the cache from source data.
   * @param headers - Array of column header names
   * @param data - 2D array of data rows (excluding headers)
   */
  buildFromData(headers: string[], data: CellValue[][]): void {
    this._recordCount = data.length;

    // Initialize fields from headers
    this._fields = headers.map((name, index) => ({
      name,
      index,
      isNumeric: true,
      isDate: false,
      sharedItems: [],
      minValue: undefined,
      maxValue: undefined,
    }));

    // Analyze data to determine field types and collect unique values
    for (const row of data) {
      for (let colIdx = 0; colIdx < row.length && colIdx < this._fields.length; colIdx++) {
        const value = row[colIdx];
        const field = this._fields[colIdx];

        if (value === null || value === undefined) {
          continue;
        }

        if (typeof value === 'string') {
          field.isNumeric = false;
          if (!field.sharedItems.includes(value)) {
            field.sharedItems.push(value);
          }
        } else if (typeof value === 'number') {
          if (field.minValue === undefined || value < field.minValue) {
            field.minValue = value;
          }
          if (field.maxValue === undefined || value > field.maxValue) {
            field.maxValue = value;
          }
        } else if (value instanceof Date) {
          field.isDate = true;
          field.isNumeric = false;
        } else if (typeof value === 'boolean') {
          field.isNumeric = false;
        }
      }
    }

    // Enable date grouping flag if any date field exists
    this._dateGrouping = this._fields.some((field) => field.isDate);

    // Store records
    this._records = data;
  }

  /**
   * Get field by name
   */
  getField(name: string): PivotCacheField | undefined {
    return this._fields.find((f) => f.name === name);
  }

  /**
   * Get field index by name
   */
  getFieldIndex(name: string): number {
    const field = this._fields.find((f) => f.name === name);
    return field ? field.index : -1;
  }

  /**
   * Generate the pivotCacheDefinition XML
   */
  toDefinitionXml(recordsRelId: string): string {
    const cacheFieldNodes: XmlNode[] = this._fields.map((field) => {
      const sharedItemsAttrs: Record<string, string> = {};
      const sharedItemChildren: XmlNode[] = [];

      if (field.sharedItems.length > 0) {
        // String field with shared items - Excel just uses count attribute
        sharedItemsAttrs.count = String(field.sharedItems.length);

        for (const item of field.sharedItems) {
          sharedItemChildren.push(createElement('s', { v: item }, []));
        }
      } else if (field.isDate) {
        sharedItemsAttrs.containsDate = '1';
      } else if (field.isNumeric) {
        // Numeric field - use "0"/"1" for boolean attributes as Excel expects
        sharedItemsAttrs.containsSemiMixedTypes = '0';
        sharedItemsAttrs.containsString = '0';
        sharedItemsAttrs.containsNumber = '1';
        // Check if all values are integers
        if (field.minValue !== undefined && field.maxValue !== undefined) {
          const isInteger = Number.isInteger(field.minValue) && Number.isInteger(field.maxValue);
          if (isInteger) {
            sharedItemsAttrs.containsInteger = '1';
          }
          sharedItemsAttrs.minValue = String(field.minValue);
          sharedItemsAttrs.maxValue = String(field.maxValue);
        }
      }

      const sharedItemsNode = createElement('sharedItems', sharedItemsAttrs, sharedItemChildren);
      return createElement('cacheField', { name: field.name, numFmtId: '0' }, [sharedItemsNode]);
    });

    const cacheFieldsNode = createElement('cacheFields', { count: String(this._fields.length) }, cacheFieldNodes);

    const worksheetSourceNode = createElement(
      'worksheetSource',
      { ref: this._sourceRange, sheet: this._sourceSheet },
      [],
    );
    const cacheSourceAttrs: Record<string, string> = { type: 'worksheet' };
    if (this._dateGrouping) {
      cacheSourceAttrs.grouping = '1';
    }
    const cacheSourceNode = createElement('cacheSource', cacheSourceAttrs, [worksheetSourceNode]);

    // Build attributes - refreshOnLoad should come early per OOXML schema
    const definitionAttrs: Record<string, string> = {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:id': recordsRelId,
    };

    // Add refreshOnLoad early in attributes (default is true)
    if (this._refreshOnLoad) {
      definitionAttrs.refreshOnLoad = '1';
    }

    // Continue with remaining attributes
    definitionAttrs.refreshedBy = 'User';
    definitionAttrs.refreshedVersion = '8';
    definitionAttrs.minRefreshableVersion = '3';
    definitionAttrs.createdVersion = '8';
    definitionAttrs.recordCount = String(this._recordCount);

    const definitionNode = createElement('pivotCacheDefinition', definitionAttrs, [cacheSourceNode, cacheFieldsNode]);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([definitionNode])}`;
  }

  /**
   * Generate the pivotCacheRecords XML
   */
  toRecordsXml(): string {
    const recordNodes: XmlNode[] = [];

    for (const row of this._records) {
      const fieldNodes: XmlNode[] = [];

      for (let colIdx = 0; colIdx < this._fields.length; colIdx++) {
        const field = this._fields[colIdx];
        const value = colIdx < row.length ? row[colIdx] : null;

        if (value === null || value === undefined) {
          // Missing value
          fieldNodes.push(createElement('m', {}, []));
        } else if (typeof value === 'string') {
          // String value - use index into sharedItems
          const idx = field.sharedItems.indexOf(value);
          if (idx >= 0) {
            fieldNodes.push(createElement('x', { v: String(idx) }, []));
          } else {
            // Direct string value (shouldn't happen if cache is built correctly)
            fieldNodes.push(createElement('s', { v: value }, []));
          }
        } else if (typeof value === 'number') {
          fieldNodes.push(createElement('n', { v: String(value) }, []));
        } else if (typeof value === 'boolean') {
          fieldNodes.push(createElement('b', { v: value ? '1' : '0' }, []));
        } else if (value instanceof Date) {
          fieldNodes.push(createElement('d', { v: value.toISOString() }, []));
        } else {
          // Unknown type, treat as missing
          fieldNodes.push(createElement('m', {}, []));
        }
      }

      recordNodes.push(createElement('r', {}, fieldNodes));
    }

    const recordsNode = createElement(
      'pivotCacheRecords',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        count: String(this._recordCount),
      },
      recordNodes,
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([recordsNode])}`;
  }
}
