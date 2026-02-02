import { readFile, writeFile } from 'fs/promises';
import type {
  SheetDefinition,
  Relationship,
  PivotTableConfig,
  CellValue,
  SheetFromDataConfig,
  ColumnConfig,
  RichCellValue,
  DateHandling,
} from './types';
import { Worksheet } from './worksheet';
import { SharedStrings } from './shared-strings';
import { Styles } from './styles';
import { PivotTable } from './pivot-table';
import { PivotCache } from './pivot-cache';
import { readZip, writeZip, readZipText, writeZipText, ZipFiles } from './utils/zip';
import { parseAddress, parseRange, toAddress } from './utils/address';
import { parseXml, findElement, getChildren, getAttr, XmlNode, stringifyXml, createElement } from './utils/xml';

/**
 * Represents an Excel workbook (.xlsx file)
 */
export class Workbook {
  private _files: ZipFiles = new Map();
  private _sheets: Map<string, Worksheet> = new Map();
  private _sheetDefs: SheetDefinition[] = [];
  private _relationships: Relationship[] = [];
  private _sharedStrings: SharedStrings;
  private _styles: Styles;
  private _dirty = false;

  // Pivot table support
  private _pivotTables: PivotTable[] = [];
  private _pivotCaches: PivotCache[] = [];
  private _nextCacheId = 0;

  // Date serialization handling
  private _dateHandling: DateHandling = 'jsDate';

  private constructor() {
    this._sharedStrings = new SharedStrings();
    this._styles = Styles.createDefault();
  }

  /**
   * Load a workbook from a file path
   */
  static async fromFile(path: string): Promise<Workbook> {
    const data = await readFile(path);
    return Workbook.fromBuffer(new Uint8Array(data));
  }

  /**
   * Load a workbook from a buffer
   */
  static async fromBuffer(data: Uint8Array): Promise<Workbook> {
    const workbook = new Workbook();
    workbook._files = await readZip(data);

    // Parse workbook.xml for sheet definitions
    const workbookXml = readZipText(workbook._files, 'xl/workbook.xml');
    if (workbookXml) {
      workbook._parseWorkbook(workbookXml);
    }

    // Parse relationships
    const relsXml = readZipText(workbook._files, 'xl/_rels/workbook.xml.rels');
    if (relsXml) {
      workbook._parseRelationships(relsXml);
    }

    // Parse shared strings
    const sharedStringsXml = readZipText(workbook._files, 'xl/sharedStrings.xml');
    if (sharedStringsXml) {
      workbook._sharedStrings = SharedStrings.parse(sharedStringsXml);
    }

    // Parse styles
    const stylesXml = readZipText(workbook._files, 'xl/styles.xml');
    if (stylesXml) {
      workbook._styles = Styles.parse(stylesXml);
    }

    return workbook;
  }

  /**
   * Create a new empty workbook
   */
  static create(): Workbook {
    const workbook = new Workbook();
    workbook._dirty = true;

    return workbook;
  }

  /**
   * Get sheet names
   */
  get sheetNames(): string[] {
    return this._sheetDefs.map((s) => s.name);
  }

  /**
   * Get number of sheets
   */
  get sheetCount(): number {
    return this._sheetDefs.length;
  }

  /**
   * Get shared strings table
   */
  get sharedStrings(): SharedStrings {
    return this._sharedStrings;
  }

  /**
   * Get styles
   */
  get styles(): Styles {
    return this._styles;
  }

  /**
   * Get the workbook date handling strategy.
   */
  get dateHandling(): DateHandling {
    return this._dateHandling;
  }

  /**
   * Set the workbook date handling strategy.
   */
  set dateHandling(value: DateHandling) {
    this._dateHandling = value;
  }

  /**
   * Get a worksheet by name or index
   */
  sheet(nameOrIndex: string | number): Worksheet {
    let def: SheetDefinition | undefined;

    if (typeof nameOrIndex === 'number') {
      def = this._sheetDefs[nameOrIndex];
    } else {
      def = this._sheetDefs.find((s) => s.name === nameOrIndex);
    }

    if (!def) {
      throw new Error(`Sheet not found: ${nameOrIndex}`);
    }

    // Return cached worksheet if available
    if (this._sheets.has(def.name)) {
      return this._sheets.get(def.name)!;
    }

    // Load worksheet
    const worksheet = new Worksheet(this, def.name);

    // Find the relationship to get the file path
    const rel = this._relationships.find((r) => r.id === def.rId);
    if (rel) {
      const sheetPath = `xl/${rel.target}`;
      const sheetXml = readZipText(this._files, sheetPath);
      if (sheetXml) {
        worksheet.parse(sheetXml);
      }
    }

    this._sheets.set(def.name, worksheet);
    return worksheet;
  }

  /**
   * Add a new worksheet
   */
  addSheet(name: string, index?: number): Worksheet {
    // Check for duplicate name
    if (this._sheetDefs.some((s) => s.name === name)) {
      throw new Error(`Sheet already exists: ${name}`);
    }

    this._dirty = true;

    // Generate new sheet ID and relationship ID
    const sheetId = Math.max(0, ...this._sheetDefs.map((s) => s.sheetId)) + 1;
    const rId = `rId${Math.max(0, ...this._relationships.map((r) => parseInt(r.id.replace('rId', ''), 10) || 0)) + 1}`;

    const def: SheetDefinition = { name, sheetId, rId };

    // Add relationship
    this._relationships.push({
      id: rId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      target: `worksheets/sheet${sheetId}.xml`,
    });

    // Insert at index or append
    if (index !== undefined && index >= 0 && index < this._sheetDefs.length) {
      this._sheetDefs.splice(index, 0, def);
    } else {
      this._sheetDefs.push(def);
    }

    // Create worksheet
    const worksheet = new Worksheet(this, name);
    this._sheets.set(name, worksheet);

    return worksheet;
  }

  /**
   * Delete a worksheet by name or index
   */
  deleteSheet(nameOrIndex: string | number): void {
    let index: number;

    if (typeof nameOrIndex === 'number') {
      index = nameOrIndex;
    } else {
      index = this._sheetDefs.findIndex((s) => s.name === nameOrIndex);
    }

    if (index < 0 || index >= this._sheetDefs.length) {
      throw new Error(`Sheet not found: ${nameOrIndex}`);
    }

    if (this._sheetDefs.length === 1) {
      throw new Error('Cannot delete the last sheet');
    }

    this._dirty = true;

    const def = this._sheetDefs[index];
    this._sheetDefs.splice(index, 1);
    this._sheets.delete(def.name);

    // Remove relationship
    const relIndex = this._relationships.findIndex((r) => r.id === def.rId);
    if (relIndex >= 0) {
      this._relationships.splice(relIndex, 1);
    }
  }

  /**
   * Rename a worksheet
   */
  renameSheet(oldName: string, newName: string): void {
    const def = this._sheetDefs.find((s) => s.name === oldName);
    if (!def) {
      throw new Error(`Sheet not found: ${oldName}`);
    }

    if (this._sheetDefs.some((s) => s.name === newName)) {
      throw new Error(`Sheet already exists: ${newName}`);
    }

    this._dirty = true;

    // Update cached worksheet
    const worksheet = this._sheets.get(oldName);
    if (worksheet) {
      worksheet.name = newName;
      this._sheets.delete(oldName);
      this._sheets.set(newName, worksheet);
    }

    def.name = newName;
  }

  /**
   * Copy a worksheet
   */
  copySheet(sourceName: string, newName: string): Worksheet {
    const source = this.sheet(sourceName);
    const copy = this.addSheet(newName);

    // Copy all cells
    for (const [address, cell] of source.cells) {
      const newCell = copy.cell(address);
      newCell.value = cell.value;
      if (cell.formula) {
        newCell.formula = cell.formula;
      }
      if (cell.styleIndex !== undefined) {
        newCell.styleIndex = cell.styleIndex;
      }
    }

    // Copy merged cells
    for (const mergedRange of source.mergedCells) {
      copy.mergeCells(mergedRange);
    }

    return copy;
  }

  /**
   * Create a new worksheet from an array of objects.
   *
   * The first row contains headers (object keys or custom column headers),
   * and subsequent rows contain the object values.
   *
   * @param config - Configuration for the sheet creation
   * @returns The created Worksheet
   *
   * @example
   * ```typescript
   * const data = [
   *   { name: 'Alice', age: 30, city: 'Paris' },
   *   { name: 'Bob', age: 25, city: 'London' },
   *   { name: 'Charlie', age: 35, city: 'Berlin' },
   * ];
   *
   * // Simple usage - all object keys become columns
   * const sheet = wb.addSheetFromData({
   *   name: 'People',
   *   data: data,
   * });
   *
   * // With custom column configuration
   * const sheet2 = wb.addSheetFromData({
   *   name: 'People Custom',
   *   data: data,
   *   columns: [
   *     { key: 'name', header: 'Full Name' },
   *     { key: 'age', header: 'Age (years)' },
   *   ],
   * });
   *
   * // With rich cell values (value, formula, style)
   * const dataWithFormulas = [
   *   { product: 'Widget', price: 10, qty: 5, total: { formula: 'B2*C2', style: { bold: true } } },
   *   { product: 'Gadget', price: 20, qty: 3, total: { formula: 'B3*C3', style: { bold: true } } },
   * ];
   * const sheet3 = wb.addSheetFromData({
   *   name: 'With Formulas',
   *   data: dataWithFormulas,
   * });
   * ```
   */
  addSheetFromData<T extends object>(config: SheetFromDataConfig<T>): Worksheet {
    const { name, data, columns, headerStyle = true, startCell = 'A1' } = config;

    if (!data?.length) return this.addSheet(name);

    // Create the new sheet
    const sheet = this.addSheet(name);

    // Parse start cell
    const startAddr = parseAddress(startCell);
    let startRow = startAddr.row;
    const startCol = startAddr.col;

    // Determine columns to use
    const columnConfigs: ColumnConfig<T>[] = columns ?? this._inferColumns(data[0]);

    // Write header row
    for (let colIdx = 0; colIdx < columnConfigs.length; colIdx++) {
      const colConfig = columnConfigs[colIdx];
      const headerText = colConfig.header ?? String(colConfig.key);
      const cell = sheet.cell(startRow, startCol + colIdx);
      cell.value = headerText;

      // Apply header style if enabled
      if (headerStyle) {
        cell.style = { bold: true };
      }
    }

    // Move to data rows
    startRow++;

    // Write data rows
    for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
      const rowData = data[rowIdx];

      for (let colIdx = 0; colIdx < columnConfigs.length; colIdx++) {
        const colConfig = columnConfigs[colIdx];
        const value = rowData[colConfig.key];
        const cell = sheet.cell(startRow + rowIdx, startCol + colIdx);

        // Check if value is a rich cell definition
        if (this._isRichCellValue(value)) {
          const richValue = value as RichCellValue;
          if (richValue.value !== undefined) cell.value = richValue.value;
          if (richValue.formula !== undefined) cell.formula = richValue.formula;
          if (richValue.style !== undefined) cell.style = richValue.style;
        } else {
          // Convert value to CellValue
          cell.value = this._toCellValue(value);
        }

        // Apply column style if defined (merged with cell style)
        if (colConfig.style) {
          cell.style = { ...cell.style, ...colConfig.style };
        }
      }
    }

    return sheet;
  }

  /**
   * Check if a value is a rich cell value object with value, formula, or style fields
   */
  private _isRichCellValue(value: unknown): value is RichCellValue {
    if (value === null || value === undefined) {
      return false;
    }
    if (typeof value !== 'object' || value instanceof Date) {
      return false;
    }
    // Check if it has at least one of the rich cell properties
    const obj = value as Record<string, unknown>;
    return 'value' in obj || 'formula' in obj || 'style' in obj;
  }

  /**
   * Infer column configuration from the first data object
   */
  private _inferColumns<T extends object>(sample: T): ColumnConfig<T>[] {
    return (Object.keys(sample) as (keyof T)[]).map((key) => ({
      key,
    }));
  }

  /**
   * Convert an unknown value to a CellValue
   */
  private _toCellValue(value: unknown): CellValue {
    if (value === null || value === undefined) {
      return null;
    }
    if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
      return value;
    }
    if (value instanceof Date) {
      return value;
    }
    if (typeof value === 'object' && 'error' in value) {
      return value as CellValue;
    }
    // Convert other types to string
    return String(value);
  }

  /**
   * Create a pivot table from source data.
   *
   * @param config - Pivot table configuration
   * @returns PivotTable instance for fluent configuration
   *
   * @example
   * ```typescript
   * const pivot = wb.createPivotTable({
   *   name: 'SalesPivot',
   *   source: 'DataSheet!A1:D100',
   *   target: 'PivotSheet!A3',
   * });
   *
   * pivot
   *   .addRowField('Region')
   *   .addColumnField('Product')
   *   .addValueField('Sales', 'sum', 'Total Sales');
   * ```
   */
  createPivotTable(config: PivotTableConfig): PivotTable {
    this._dirty = true;

    // Parse source reference (Sheet!Range)
    const { sheetName: sourceSheet, range: sourceRange } = this._parseSheetRef(config.source);

    // Parse target reference
    const { sheetName: targetSheet, range: targetCell } = this._parseSheetRef(config.target);

    // Ensure target sheet exists
    if (!this._sheetDefs.some((s) => s.name === targetSheet)) {
      this.addSheet(targetSheet);
    }

    // Parse target cell address
    const targetAddr = parseAddress(targetCell);

    // Get source worksheet and extract data
    const sourceWs = this.sheet(sourceSheet);
    const { headers, data } = this._extractSourceData(sourceWs, sourceRange);

    // Create pivot cache
    const cacheId = this._nextCacheId++;
    const cache = new PivotCache(cacheId, sourceSheet, sourceRange);
    cache.buildFromData(headers, data);
    // refreshOnLoad defaults to true; only disable if explicitly set to false
    if (config.refreshOnLoad === false) {
      cache.refreshOnLoad = false;
    }
    this._pivotCaches.push(cache);

    // Create pivot table
    const pivotTableIndex = this._pivotTables.length + 1;
    const pivotTable = new PivotTable(
      config.name,
      cache,
      targetSheet,
      targetCell,
      targetAddr.row + 1, // Convert to 1-based
      targetAddr.col,
      pivotTableIndex,
    );

    // Set styles reference for number format resolution
    pivotTable.setStyles(this._styles);

    this._pivotTables.push(pivotTable);

    return pivotTable;
  }

  /**
   * Parse a sheet reference like "Sheet1!A1:D100" into sheet name and range
   */
  private _parseSheetRef(ref: string): { sheetName: string; range: string } {
    const match = ref.match(/^(.+?)!(.+)$/);
    if (!match) {
      throw new Error(`Invalid reference format: ${ref}. Expected "SheetName!Range"`);
    }
    return { sheetName: match[1], range: match[2] };
  }

  /**
   * Extract headers and data from a source range
   */
  private _extractSourceData(sheet: Worksheet, rangeStr: string): { headers: string[]; data: CellValue[][] } {
    const range = parseRange(rangeStr);
    const headers: string[] = [];
    const data: CellValue[][] = [];

    // First row is headers
    for (let col = range.start.col; col <= range.end.col; col++) {
      const cell = sheet.cell(toAddress(range.start.row, col));
      headers.push(String(cell.value ?? `Column${col + 1}`));
    }

    // Remaining rows are data
    for (let row = range.start.row + 1; row <= range.end.row; row++) {
      const rowData: CellValue[] = [];
      for (let col = range.start.col; col <= range.end.col; col++) {
        const cell = sheet.cell(toAddress(row, col));
        rowData.push(cell.value);
      }
      data.push(rowData);
    }

    return { headers, data };
  }

  /**
   * Save the workbook to a file
   */
  async toFile(path: string): Promise<void> {
    const buffer = await this.toBuffer();
    await writeFile(path, buffer);
  }

  /**
   * Save the workbook to a buffer
   */
  async toBuffer(): Promise<Uint8Array> {
    // Update files map with modified content
    this._updateFiles();

    // Write ZIP
    return writeZip(this._files);
  }

  private _parseWorkbook(xml: string): void {
    const parsed = parseXml(xml);
    const workbook = findElement(parsed, 'workbook');
    if (!workbook) return;

    const children = getChildren(workbook, 'workbook');
    const sheets = findElement(children, 'sheets');
    if (!sheets) return;

    for (const child of getChildren(sheets, 'sheets')) {
      if ('sheet' in child) {
        const name = getAttr(child, 'name');
        const sheetId = getAttr(child, 'sheetId');
        const rId = getAttr(child, 'r:id');

        if (name && sheetId && rId) {
          this._sheetDefs.push({
            name,
            sheetId: parseInt(sheetId, 10),
            rId,
          });
        }
      }
    }
  }

  private _parseRelationships(xml: string): void {
    const parsed = parseXml(xml);
    const rels = findElement(parsed, 'Relationships');
    if (!rels) return;

    for (const child of getChildren(rels, 'Relationships')) {
      if ('Relationship' in child) {
        const id = getAttr(child, 'Id');
        const type = getAttr(child, 'Type');
        const target = getAttr(child, 'Target');

        if (id && type && target) {
          this._relationships.push({ id, type, target });
        }
      }
    }
  }

  private _updateFiles(): void {
    // Update workbook.xml
    this._updateWorkbookXml();

    // Update relationships
    this._updateRelationshipsXml();

    // Update content types
    this._updateContentTypes();

    // Update shared strings if modified
    if (this._sharedStrings.dirty || this._sharedStrings.count > 0) {
      writeZipText(this._files, 'xl/sharedStrings.xml', this._sharedStrings.toXml());
    }

    // Update styles if modified or if file doesn't exist yet
    if (this._styles.dirty || this._dirty || !this._files.has('xl/styles.xml')) {
      writeZipText(this._files, 'xl/styles.xml', this._styles.toXml());
    }

    // Update worksheets
    for (const [name, worksheet] of this._sheets) {
      if (worksheet.dirty || this._dirty) {
        const def = this._sheetDefs.find((s) => s.name === name);
        if (def) {
          const rel = this._relationships.find((r) => r.id === def.rId);
          if (rel) {
            const sheetPath = `xl/${rel.target}`;
            writeZipText(this._files, sheetPath, worksheet.toXml());
          }
        }
      }
    }

    // Update pivot tables
    if (this._pivotTables.length > 0) {
      this._updatePivotTableFiles();
    }
  }

  private _updateWorkbookXml(): void {
    const sheetNodes: XmlNode[] = this._sheetDefs.map((def) =>
      createElement('sheet', { name: def.name, sheetId: String(def.sheetId), 'r:id': def.rId }, []),
    );

    const sheetsNode = createElement('sheets', {}, sheetNodes);

    const children: XmlNode[] = [sheetsNode];

    // Add pivot caches if any
    if (this._pivotCaches.length > 0) {
      const pivotCacheNodes: XmlNode[] = this._pivotCaches.map((cache, idx) => {
        // Cache relationship ID is after sheets, sharedStrings, and styles
        const cacheRelId = `rId${this._relationships.length + 3 + idx}`;
        return createElement('pivotCache', { cacheId: String(cache.cacheId), 'r:id': cacheRelId }, []);
      });
      children.push(createElement('pivotCaches', {}, pivotCacheNodes));
    }

    const workbookNode = createElement(
      'workbook',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      },
      children,
    );

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([workbookNode])}`;
    writeZipText(this._files, 'xl/workbook.xml', xml);
  }

  private _updateRelationshipsXml(): void {
    const relNodes: XmlNode[] = this._relationships.map((rel) =>
      createElement('Relationship', { Id: rel.id, Type: rel.type, Target: rel.target }, []),
    );

    // Calculate next available relationship ID based on existing max ID
    let nextRelId = Math.max(0, ...this._relationships.map((r) => parseInt(r.id.replace('rId', ''), 10) || 0)) + 1;

    // Add shared strings relationship if needed
    if (this._sharedStrings.count > 0) {
      const hasSharedStrings = this._relationships.some(
        (r) => r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
      );
      if (!hasSharedStrings) {
        relNodes.push(
          createElement(
            'Relationship',
            {
              Id: `rId${nextRelId++}`,
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
              Target: 'sharedStrings.xml',
            },
            [],
          ),
        );
      }
    }

    // Add styles relationship if needed
    const hasStyles = this._relationships.some(
      (r) => r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    );
    if (!hasStyles) {
      relNodes.push(
        createElement(
          'Relationship',
          {
            Id: `rId${nextRelId++}`,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            Target: 'styles.xml',
          },
          [],
        ),
      );
    }

    // Add pivot cache relationships
    for (let i = 0; i < this._pivotCaches.length; i++) {
      relNodes.push(
        createElement(
          'Relationship',
          {
            Id: `rId${nextRelId++}`,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
            Target: `pivotCache/pivotCacheDefinition${i + 1}.xml`,
          },
          [],
        ),
      );
    }

    const relsNode = createElement(
      'Relationships',
      { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
      relNodes,
    );

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([relsNode])}`;
    writeZipText(this._files, 'xl/_rels/workbook.xml.rels', xml);
  }

  private _updateContentTypes(): void {
    const types: XmlNode[] = [
      createElement(
        'Default',
        { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' },
        [],
      ),
      createElement('Default', { Extension: 'xml', ContentType: 'application/xml' }, []),
      createElement(
        'Override',
        {
          PartName: '/xl/workbook.xml',
          ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        },
        [],
      ),
      createElement(
        'Override',
        {
          PartName: '/xl/styles.xml',
          ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
        },
        [],
      ),
    ];

    // Add shared strings if present
    if (this._sharedStrings.count > 0) {
      types.push(
        createElement(
          'Override',
          {
            PartName: '/xl/sharedStrings.xml',
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          },
          [],
        ),
      );
    }

    // Add worksheets
    for (const def of this._sheetDefs) {
      const rel = this._relationships.find((r) => r.id === def.rId);
      if (rel) {
        types.push(
          createElement(
            'Override',
            {
              PartName: `/xl/${rel.target}`,
              ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            },
            [],
          ),
        );
      }
    }

    // Add pivot cache definitions and records
    for (let i = 0; i < this._pivotCaches.length; i++) {
      types.push(
        createElement(
          'Override',
          {
            PartName: `/xl/pivotCache/pivotCacheDefinition${i + 1}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
          },
          [],
        ),
      );
      types.push(
        createElement(
          'Override',
          {
            PartName: `/xl/pivotCache/pivotCacheRecords${i + 1}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
          },
          [],
        ),
      );
    }

    // Add pivot tables
    for (let i = 0; i < this._pivotTables.length; i++) {
      types.push(
        createElement(
          'Override',
          {
            PartName: `/xl/pivotTables/pivotTable${i + 1}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
          },
          [],
        ),
      );
    }

    const typesNode = createElement(
      'Types',
      { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' },
      types,
    );

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([typesNode])}`;
    writeZipText(this._files, '[Content_Types].xml', xml);

    // Also ensure _rels/.rels exists
    const rootRelsXml = readZipText(this._files, '_rels/.rels');
    if (!rootRelsXml) {
      const rootRels = createElement(
        'Relationships',
        { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
        [
          createElement(
            'Relationship',
            {
              Id: 'rId1',
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
              Target: 'xl/workbook.xml',
            },
            [],
          ),
        ],
      );
      writeZipText(
        this._files,
        '_rels/.rels',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([rootRels])}`,
      );
    }
  }

  /**
   * Generate all pivot table related files
   */
  private _updatePivotTableFiles(): void {
    // Track which sheets have pivot tables for their .rels files
    const sheetPivotTables: Map<string, PivotTable[]> = new Map();

    for (const pivotTable of this._pivotTables) {
      const sheetName = pivotTable.targetSheet;
      if (!sheetPivotTables.has(sheetName)) {
        sheetPivotTables.set(sheetName, []);
      }
      sheetPivotTables.get(sheetName)!.push(pivotTable);
    }

    // Generate pivot cache files
    for (let i = 0; i < this._pivotCaches.length; i++) {
      const cache = this._pivotCaches[i];
      const cacheIdx = i + 1;

      // Pivot cache definition
      const definitionPath = `xl/pivotCache/pivotCacheDefinition${cacheIdx}.xml`;
      writeZipText(this._files, definitionPath, cache.toDefinitionXml('rId1'));

      // Pivot cache records
      const recordsPath = `xl/pivotCache/pivotCacheRecords${cacheIdx}.xml`;
      writeZipText(this._files, recordsPath, cache.toRecordsXml());

      // Pivot cache definition relationships (link to records)
      const cacheRelsPath = `xl/pivotCache/_rels/pivotCacheDefinition${cacheIdx}.xml.rels`;
      const cacheRels = createElement(
        'Relationships',
        { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
        [
          createElement(
            'Relationship',
            {
              Id: 'rId1',
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
              Target: `pivotCacheRecords${cacheIdx}.xml`,
            },
            [],
          ),
        ],
      );
      writeZipText(
        this._files,
        cacheRelsPath,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([cacheRels])}`,
      );
    }

    // Generate pivot table files
    for (let i = 0; i < this._pivotTables.length; i++) {
      const pivotTable = this._pivotTables[i];
      const ptIdx = i + 1;

      // Pivot table definition
      const ptPath = `xl/pivotTables/pivotTable${ptIdx}.xml`;
      writeZipText(this._files, ptPath, pivotTable.toXml());

      // Pivot table relationships (link to cache definition)
      const cacheIdx = this._pivotCaches.indexOf(pivotTable.cache) + 1;
      const ptRelsPath = `xl/pivotTables/_rels/pivotTable${ptIdx}.xml.rels`;
      const ptRels = createElement(
        'Relationships',
        { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
        [
          createElement(
            'Relationship',
            {
              Id: 'rId1',
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
              Target: `../pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
            },
            [],
          ),
        ],
      );
      writeZipText(
        this._files,
        ptRelsPath,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([ptRels])}`,
      );
    }

    // Generate worksheet relationships for pivot tables
    for (const [sheetName, pivotTables] of sheetPivotTables) {
      const def = this._sheetDefs.find((s) => s.name === sheetName);
      if (!def) continue;

      const rel = this._relationships.find((r) => r.id === def.rId);
      if (!rel) continue;

      // Extract sheet file name from target path
      const sheetFileName = rel.target.split('/').pop();
      const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;

      const relNodes: XmlNode[] = [];
      for (let i = 0; i < pivotTables.length; i++) {
        const pt = pivotTables[i];
        relNodes.push(
          createElement(
            'Relationship',
            {
              Id: `rId${i + 1}`,
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
              Target: `../pivotTables/pivotTable${pt.index}.xml`,
            },
            [],
          ),
        );
      }

      const sheetRels = createElement(
        'Relationships',
        { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
        relNodes,
      );
      writeZipText(
        this._files,
        sheetRelsPath,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([sheetRels])}`,
      );
    }
  }
}
