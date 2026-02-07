import type { PivotCacheField, CellValue } from './types';
import type { Styles } from './styles';
import { createElement, stringifyXml, XmlNode } from './utils/xml';

/**
 * Manages the pivot cache (definition and records) for a pivot table.
 * The cache stores source data metadata and cached values.
 */
export class PivotCache {
  private _cacheId: number;
  private _fileIndex: number;
  private _sourceSheet: string;
  private _sourceRange: string;
  private _fields: PivotCacheField[] = [];
  private _records: CellValue[][] = [];
  private _recordCount = 0;
  private _saveData = true;
  private _refreshOnLoad = true; // Default to true
  // Optimized lookup: Map<fieldIndex, Map<stringValue, sharedItemsIndex>>
  private _sharedItemsIndexMap: Map<number, Map<string, number>> = new Map();
  private _blankItemIndexMap: Map<number, number> = new Map();
  private _styles: Styles | null = null;

  constructor(cacheId: number, sourceSheet: string, sourceRange: string, fileIndex: number) {
    this._cacheId = cacheId;
    this._fileIndex = fileIndex;
    this._sourceSheet = sourceSheet;
    this._sourceRange = sourceRange;
  }

  /**
   * Set styles reference for number format resolution.
   * @internal
   */
  setStyles(styles: Styles): void {
    this._styles = styles;
  }

  /**
   * Get the cache ID
   */
  get cacheId(): number {
    return this._cacheId;
  }

  /**
   * Get the file index for this cache (used for file naming).
   */
  get fileIndex(): number {
    return this._fileIndex;
  }

  /**
   * Set refreshOnLoad option
   */
  set refreshOnLoad(value: boolean) {
    this._refreshOnLoad = value;
  }

  /**
   * Set saveData option
   */
  set saveData(value: boolean) {
    this._saveData = value;
  }

  /**
   * Get refreshOnLoad option
   */
  get refreshOnLoad(): boolean {
    return this._refreshOnLoad;
  }

  /**
   * Get saveData option
   */
  get saveData(): boolean {
    return this._saveData;
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
      hasBoolean: false,
      hasBlank: false,
      numFmtId: undefined,
      sharedItems: [],
      minValue: undefined,
      maxValue: undefined,
      minDate: undefined,
      maxDate: undefined,
    }));

    // Use Maps for unique value collection during analysis
    const sharedItemsMaps: Map<string, string>[] = this._fields.map(() => new Map<string, string>());

    // Analyze data to determine field types and collect unique values
    for (const row of data) {
      for (let colIdx = 0; colIdx < row.length && colIdx < this._fields.length; colIdx++) {
        const value = row[colIdx];
        const field = this._fields[colIdx];

        if (value === null || value === undefined) {
          field.hasBlank = true;
          continue;
        }

        if (typeof value === 'string') {
          field.isNumeric = false;
          const map = sharedItemsMaps[colIdx];
          if (!map.has(value)) {
            map.set(value, value);
          }
        } else if (typeof value === 'number') {
          if (field.isDate) {
            const d = this._excelSerialToDate(value);
            if (!field.minDate || d < field.minDate) {
              field.minDate = d;
            }
            if (!field.maxDate || d > field.maxDate) {
              field.maxDate = d;
            }
          } else {
            if (field.minValue === undefined || value < field.minValue) {
              field.minValue = value;
            }
            if (field.maxValue === undefined || value > field.maxValue) {
              field.maxValue = value;
            }
          }
        } else if (value instanceof Date) {
          field.isDate = true;
          field.isNumeric = false;
          if (!field.minDate || value < field.minDate) {
            field.minDate = value;
          }
          if (!field.maxDate || value > field.maxDate) {
            field.maxDate = value;
          }
        } else if (typeof value === 'boolean') {
          field.isNumeric = false;
          field.hasBoolean = true;
        }
      }
    }

    // Resolve number formats if styles are available
    if (this._styles) {
      const dateFmtId = this._styles.getOrCreateNumFmtId('mm-dd-yy');
      for (const field of this._fields) {
        if (field.isDate) {
          field.numFmtId = dateFmtId;
        }
      }
    }

    // Convert Sets to arrays and build reverse index Maps for O(1) lookup during XML generation
    this._sharedItemsIndexMap.clear();
    this._blankItemIndexMap.clear();
    for (let colIdx = 0; colIdx < this._fields.length; colIdx++) {
      const field = this._fields[colIdx];
      const map = sharedItemsMaps[colIdx];

      // Convert Map values to array (maintains insertion order in ES6+)
      field.sharedItems = Array.from(map.values());

      // Build reverse lookup Map: value -> index
      if (field.sharedItems.length > 0) {
        const indexMap = new Map<string, number>();
        for (let i = 0; i < field.sharedItems.length; i++) {
          indexMap.set(field.sharedItems[i], i);
        }
        this._sharedItemsIndexMap.set(colIdx, indexMap);

        if (field.hasBlank) {
          const blankIndex = field.sharedItems.length;
          this._blankItemIndexMap.set(colIdx, blankIndex);
        }
      }
    }

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
        // String field with shared items
        const total = field.hasBlank ? field.sharedItems.length + 1 : field.sharedItems.length;
        sharedItemsAttrs.count = String(total);
        sharedItemsAttrs.containsString = '1';

        if (field.hasBlank) {
          sharedItemsAttrs.containsBlank = '1';
        }

        for (const item of field.sharedItems) {
          sharedItemChildren.push(createElement('s', { v: item }, []));
        }
        if (field.hasBlank) {
          sharedItemChildren.push(createElement('m', {}, []));
        }
      } else if (field.isDate) {
        sharedItemsAttrs.containsSemiMixedTypes = '0';
        sharedItemsAttrs.containsString = '0';
        sharedItemsAttrs.containsDate = '1';
        sharedItemsAttrs.containsNonDate = '0';
        if (field.hasBlank) {
          sharedItemsAttrs.containsBlank = '1';
        }
        if (field.minDate) {
          sharedItemsAttrs.minDate = this._formatDate(field.minDate);
        }
        if (field.maxDate) {
          const maxDate = new Date(field.maxDate.getTime() + 24 * 60 * 60 * 1000);
          sharedItemsAttrs.maxDate = this._formatDate(maxDate);
        }
      } else if (field.isNumeric) {
        // Numeric field - use "0"/"1" for boolean attributes as Excel expects
        sharedItemsAttrs.containsSemiMixedTypes = '0';
        sharedItemsAttrs.containsString = '0';
        sharedItemsAttrs.containsNumber = '1';
        if (field.hasBlank) {
          sharedItemsAttrs.containsBlank = '1';
        }
        // Check if all values are integers
        if (field.minValue !== undefined && field.maxValue !== undefined) {
          const isInteger = Number.isInteger(field.minValue) && Number.isInteger(field.maxValue);
          if (isInteger) {
            sharedItemsAttrs.containsInteger = '1';
          }
          sharedItemsAttrs.minValue = this._formatNumber(field.minValue);
          sharedItemsAttrs.maxValue = this._formatNumber(field.maxValue);
        }
      } else if (field.hasBoolean) {
        // Boolean-only field (no strings, no numbers)
        if (field.hasBlank) {
          sharedItemsAttrs.containsBlank = '1';
        }
        sharedItemsAttrs.count = field.hasBlank ? '3' : '2';
        sharedItemChildren.push(createElement('b', { v: '0' }, []));
        sharedItemChildren.push(createElement('b', { v: '1' }, []));
        if (field.hasBlank) {
          sharedItemChildren.push(createElement('m', {}, []));
        }
      } else if (field.hasBlank) {
        // Field that only contains blanks
        sharedItemsAttrs.containsBlank = '1';
      }

      const sharedItemsNode = createElement('sharedItems', sharedItemsAttrs, sharedItemChildren);
      const cacheFieldAttrs: Record<string, string> = { name: field.name, numFmtId: String(field.numFmtId ?? 0) };
      return createElement('cacheField', cacheFieldAttrs, [sharedItemsNode]);
    });

    const cacheFieldsNode = createElement('cacheFields', { count: String(this._fields.length) }, cacheFieldNodes);

    const worksheetSourceNode = createElement(
      'worksheetSource',
      { ref: this._sourceRange, sheet: this._sourceSheet },
      [],
    );
    const cacheSourceNode = createElement('cacheSource', { type: 'worksheet' }, [worksheetSourceNode]);

    // Build attributes - align with Excel expectations
    const definitionAttrs: Record<string, string> = {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:id': recordsRelId,
    };

    if (this._refreshOnLoad) {
      definitionAttrs.refreshOnLoad = '1';
    }

    definitionAttrs.refreshedBy = 'User';
    definitionAttrs.refreshedVersion = '8';
    definitionAttrs.minRefreshableVersion = '3';
    definitionAttrs.createdVersion = '8';
    if (!this._saveData) {
      definitionAttrs.saveData = '0';
      definitionAttrs.recordCount = '0';
    } else {
      definitionAttrs.recordCount = String(this._recordCount);
    }

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
        const value = colIdx < row.length ? row[colIdx] : null;

        if (value === null || value === undefined) {
          // Missing value
          const blankIndex = this._blankItemIndexMap.get(colIdx);
          if (blankIndex !== undefined) {
            fieldNodes.push(createElement('x', { v: String(blankIndex) }, []));
          } else {
            fieldNodes.push(createElement('m', {}, []));
          }
        } else if (typeof value === 'string') {
          // String value - use index into sharedItems via O(1) Map lookup
          const indexMap = this._sharedItemsIndexMap.get(colIdx);
          const idx = indexMap?.get(value);
          if (idx !== undefined) {
            fieldNodes.push(createElement('x', { v: String(idx) }, []));
          } else {
            // Direct string value (shouldn't happen if cache is built correctly)
            fieldNodes.push(createElement('s', { v: value }, []));
          }
        } else if (typeof value === 'number') {
          if (this._fields[colIdx]?.isDate) {
            const d = this._excelSerialToDate(value);
            fieldNodes.push(createElement('d', { v: this._formatDate(d) }, []));
          } else {
            fieldNodes.push(createElement('n', { v: String(value) }, []));
          }
        } else if (typeof value === 'boolean') {
          fieldNodes.push(createElement('b', { v: value ? '1' : '0' }, []));
        } else if (value instanceof Date) {
          fieldNodes.push(createElement('d', { v: this._formatDate(value) }, []));
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

  private _formatDate(value: Date): string {
    return value.toISOString().replace(/\.\d{3}Z$/, '');
  }

  private _formatNumber(value: number): string {
    if (Number.isInteger(value)) {
      return String(value);
    }
    if (Math.abs(value) >= 1000000) {
      return value.toFixed(16).replace(/0+$/, '').replace(/\.$/, '');
    }
    return String(value);
  }

  private _excelSerialToDate(serial: number): Date {
    // Excel epoch: December 31, 1899
    const EXCEL_EPOCH = Date.UTC(1899, 11, 31);
    const MS_PER_DAY = 24 * 60 * 60 * 1000;
    const adjusted = serial >= 60 ? serial - 1 : serial;
    const ms = Math.round(adjusted * MS_PER_DAY);
    return new Date(EXCEL_EPOCH + ms);
  }
}
