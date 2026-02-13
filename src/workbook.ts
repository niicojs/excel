import { readFile, writeFile } from 'fs/promises';
import type {
  SheetDefinition,
  Relationship,
  CellValue,
  SheetFromDataConfig,
  ColumnConfig,
  RichCellValue,
  DateHandling,
  PivotTableConfig,
  RangeAddress,
} from './types';
import { Worksheet } from './worksheet';
import { SharedStrings } from './shared-strings';
import { Styles } from './styles';
import { PivotTable } from './pivot-table';
import { readZip, writeZip, readZipText, writeZipText, ZipFiles } from './utils/zip';
import { parseAddress, parseSheetAddress, parseSheetRange } from './utils/address';
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

  // Table support
  private _nextTableId = 1;

  // Pivot table support
  private _pivotTables: PivotTable[] = [];
  private _nextPivotTableId = 1;
  private _nextPivotCacheId = 1;

  // Date serialization handling
  private _dateHandling: DateHandling = 'jsDate';

  private _locale = 'fr-FR';

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
   * Get the workbook locale for formatting.
   */
  get locale(): string {
    return this._locale;
  }

  /**
   * Set the workbook locale for formatting.
   */
  set locale(value: string) {
    this._locale = value;
  }

  /**
   * Get the next unique table ID for this workbook.
   * Table IDs must be unique across all worksheets.
   * @internal
   */
  getNextTableId(): number {
    return this._nextTableId++;
  }

  /**
   * Get all pivot tables in the workbook.
   */
  get pivotTables(): PivotTable[] {
    return [...this._pivotTables];
  }

  /**
   * Create a new pivot table.
   */
  createPivotTable(config: PivotTableConfig): PivotTable {
    if (!config.name || config.name.trim().length === 0) {
      throw new Error('Pivot table name is required');
    }

    if (this._pivotTables.some((pivot) => pivot.name === config.name)) {
      throw new Error(`Pivot table name already exists: ${config.name}`);
    }

    const sourceRef = parseSheetRange(config.source);
    const targetRef = parseSheetAddress(config.target);

    const sourceSheet = this.sheet(sourceRef.sheet);
    this.sheet(targetRef.sheet);

    const sourceRange = this._normalizeRange(sourceRef.range);
    if (sourceRange.start.row >= sourceRange.end.row) {
      throw new Error('Pivot source range must include a header row and at least one data row');
    }

    const fields = this._extractPivotFields(sourceSheet, sourceRange);

    const cacheId = this._nextPivotCacheId++;
    const pivotId = this._nextPivotTableId++;
    const cachePartIndex = this._pivotTables.length + 1;

    const pivot = new PivotTable(
      this,
      config,
      sourceRef.sheet,
      sourceSheet,
      sourceRange,
      targetRef.sheet,
      targetRef.address,
      cacheId,
      pivotId,
      cachePartIndex,
      fields,
    );

    this._pivotTables.push(pivot);
    this._dirty = true;

    return pivot;
  }

  private _extractPivotFields(
    sourceSheet: Worksheet,
    sourceRange: RangeAddress,
  ): { name: string; sourceCol: number }[] {
    const fields: { name: string; sourceCol: number }[] = [];
    const seen = new Set<string>();

    for (let col = sourceRange.start.col; col <= sourceRange.end.col; col++) {
      const headerCell = sourceSheet.getCellIfExists(sourceRange.start.row, col);
      const rawHeader = headerCell?.value;
      const name = rawHeader == null ? `Column${col - sourceRange.start.col + 1}` : String(rawHeader).trim();

      if (!name) {
        throw new Error(`Pivot source header is empty at column ${col + 1}`);
      }

      if (seen.has(name)) {
        throw new Error(`Duplicate pivot source header: ${name}`);
      }

      seen.add(name);
      fields.push({ name, sourceCol: col });
    }

    return fields;
  }

  private _normalizeRange(range: RangeAddress): RangeAddress {
    return {
      start: {
        row: Math.min(range.start.row, range.end.row),
        col: Math.min(range.start.col, range.end.col),
      },
      end: {
        row: Math.max(range.start.row, range.end.row),
        col: Math.max(range.start.col, range.end.col),
      },
    };
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

    const rel = this._relationships.find((r) => r.id === def.rId);
    if (rel) {
      const sheetPath = `xl/${rel.target}`;
      this._files.delete(sheetPath);
    }

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

    // Copy column widths
    for (const [col, width] of source.getColumnWidths()) {
      copy.setColumnWidth(col, width);
    }

    // Copy row heights
    for (const [row, height] of source.getRowHeights()) {
      copy.setRowHeight(row, height);
    }

    // Copy frozen panes
    const frozen = source.getFrozenPane();
    if (frozen) {
      copy.freezePane(frozen.row, frozen.col);
    }

    // Copy merged cells
    for (const mergedRange of source.mergedCells) {
      copy.mergeCells(mergedRange);
    }

    // Copy tables
    for (const table of source.tables) {
      const tableName = this._createUniqueTableName(table.name, newName);
      const newTable = copy.createTable({
        name: tableName,
        range: table.baseRange,
        totalRow: table.hasTotalRow,
        headerRow: table.hasHeaderRow,
        style: table.style,
      });

      if (!table.hasAutoFilter) {
        newTable.setAutoFilter(false);
      }

      if (table.hasTotalRow) {
        for (const columnName of table.columns) {
          const fn = table.getTotalFunction(columnName);
          if (fn) {
            newTable.setTotalFunction(columnName, fn);
          }
        }
      }
    }

    return copy;
  }

  private _createUniqueTableName(base: string, sheetName: string): string {
    const normalizedSheet = sheetName.replace(/[^A-Za-z0-9_.]/g, '_');
    const sanitizedBase = this._sanitizeTableName(`${base}_${normalizedSheet || 'Sheet'}`);
    let candidate = sanitizedBase;
    let counter = 1;

    while (this._hasTableName(candidate)) {
      candidate = `${sanitizedBase}_${counter++}`;
    }

    return candidate;
  }

  private _sanitizeTableName(name: string): string {
    let result = name.replace(/[^A-Za-z0-9_.]/g, '_');
    if (!/^[A-Za-z_]/.test(result)) {
      result = `_${result}`;
    }
    if (result.length === 0) {
      result = 'Table';
    }
    return result;
  }

  private _hasTableName(name: string): boolean {
    for (const sheetName of this.sheetNames) {
      const ws = this.sheet(sheetName);
      for (const table of ws.tables) {
        if (table.name === name) return true;
      }
    }
    return false;
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
    const relationshipInfo = this._buildRelationshipInfo();

    // Update workbook.xml
    this._updateWorkbookXml(relationshipInfo.pivotCacheRelByTarget);

    // Update relationships
    this._updateRelationshipsXml(relationshipInfo.relNodes);

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
      if (worksheet.dirty || this._dirty || worksheet.tables.length > 0) {
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

    // Update tables (sets table rel IDs for tableParts)
    this._updateTableFiles();

    // Update pivot tables (sets pivot rel IDs for pivotTableParts)
    this._updatePivotFiles();

    // Update worksheets to align tableParts with relationship IDs
    for (const [name, worksheet] of this._sheets) {
      if (worksheet.dirty || this._dirty || worksheet.tables.length > 0 || this._pivotTables.length > 0) {
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
  }

  private _updateWorkbookXml(pivotCacheRelByTarget: Map<string, string>): void {
    const sheetNodes: XmlNode[] = this._sheetDefs.map((def) =>
      createElement('sheet', { name: def.name, sheetId: String(def.sheetId), 'r:id': def.rId }, []),
    );

    const sheetsNode = createElement('sheets', {}, sheetNodes);

    const children: XmlNode[] = [sheetsNode];

    if (this._pivotTables.length > 0) {
      const pivotCacheNodes: XmlNode[] = [];
      for (const pivot of this._pivotTables) {
        const target = `pivotCache/pivotCacheDefinition${pivot.cachePartIndex}.xml`;
        const relId = pivotCacheRelByTarget.get(target);
        if (!relId) continue;
        pivotCacheNodes.push(createElement('pivotCache', { cacheId: String(pivot.cacheId), 'r:id': relId }, []));
      }

      if (pivotCacheNodes.length > 0) {
        children.push(createElement('pivotCaches', {}, pivotCacheNodes));
      }
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

  private _updateRelationshipsXml(relNodes: XmlNode[]): void {
    const relsNode = createElement(
      'Relationships',
      { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
      relNodes,
    );

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([relsNode])}`;
    writeZipText(this._files, 'xl/_rels/workbook.xml.rels', xml);
  }

  private _buildRelationshipInfo(): { relNodes: XmlNode[]; pivotCacheRelByTarget: Map<string, string> } {
    const relNodes: XmlNode[] = this._relationships.map((rel) =>
      createElement('Relationship', { Id: rel.id, Type: rel.type, Target: rel.target }, []),
    );
    const pivotCacheRelByTarget = new Map<string, string>();

    const reservedRelIds = new Set<string>(relNodes.map((node) => getAttr(node, 'Id') || '').filter(Boolean));
    let nextRelId = Math.max(0, ...this._relationships.map((r) => parseInt(r.id.replace('rId', ''), 10) || 0)) + 1;

    const allocateRelId = (): string => {
      while (reservedRelIds.has(`rId${nextRelId}`)) {
        nextRelId++;
      }
      const id = `rId${nextRelId}`;
      nextRelId++;
      reservedRelIds.add(id);
      return id;
    };

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
              Id: allocateRelId(),
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
            Id: allocateRelId(),
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            Target: 'styles.xml',
          },
          [],
        ),
      );
    }

    for (const pivot of this._pivotTables) {
      const target = `pivotCache/pivotCacheDefinition${pivot.cachePartIndex}.xml`;
      const hasPivotCacheRel = relNodes.some(
        (node) =>
          getAttr(node, 'Type') ===
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition' &&
          getAttr(node, 'Target') === target,
      );

      if (!hasPivotCacheRel) {
        const id = allocateRelId();
        pivotCacheRelByTarget.set(target, id);
        relNodes.push(
          createElement(
            'Relationship',
            {
              Id: id,
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
              Target: target,
            },
            [],
          ),
        );
      } else {
        const existing = relNodes.find(
          (node) =>
            getAttr(node, 'Type') ===
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition' &&
            getAttr(node, 'Target') === target,
        );
        const existingId = existing ? getAttr(existing, 'Id') : undefined;
        if (existingId) {
          pivotCacheRelByTarget.set(target, existingId);
        }
      }
    }

    return { relNodes, pivotCacheRelByTarget };
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

    // Add tables
    let tableIndex = 1;
    for (const def of this._sheetDefs) {
      const worksheet = this._sheets.get(def.name);
      if (worksheet) {
        for (let i = 0; i < worksheet.tables.length; i++) {
          types.push(
            createElement(
              'Override',
              {
                PartName: `/xl/tables/table${tableIndex}.xml`,
                ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
              },
              [],
            ),
          );
          tableIndex++;
        }
      }
    }

    // Add pivot caches and pivot tables
    for (const pivot of this._pivotTables) {
      types.push(
        createElement(
          'Override',
          {
            PartName: `/xl/pivotCache/pivotCacheDefinition${pivot.cachePartIndex}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
          },
          [],
        ),
      );

      types.push(
        createElement(
          'Override',
          {
            PartName: `/xl/pivotCache/pivotCacheRecords${pivot.cachePartIndex}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
          },
          [],
        ),
      );

      types.push(
        createElement(
          'Override',
          {
            PartName: `/xl/pivotTables/pivotTable${pivot.pivotId}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
          },
          [],
        ),
      );
    }

    const existingTypesXml = readZipText(this._files, '[Content_Types].xml');
    const existingKeys = new Set(
      types
        .map((t) => {
          if ('Default' in t) {
            const a = t[':@'] as Record<string, string> | undefined;
            return `Default:${a?.['@_Extension'] || ''}`;
          }
          if ('Override' in t) {
            const a = t[':@'] as Record<string, string> | undefined;
            return `Override:${a?.['@_PartName'] || ''}`;
          }
          return '';
        })
        .filter(Boolean),
    );
    if (existingTypesXml) {
      const parsed = parseXml(existingTypesXml);
      const typesElement = findElement(parsed, 'Types');
      if (typesElement) {
        const existingNodes = getChildren(typesElement, 'Types');
        for (const node of existingNodes) {
          if ('Default' in node || 'Override' in node) {
            const type = 'Default' in node ? 'Default' : 'Override';
            const attrs = node[':@'] as Record<string, string> | undefined;
            const key =
              type === 'Default'
                ? `Default:${attrs?.['@_Extension'] || ''}`
                : `Override:${attrs?.['@_PartName'] || ''}`;
            if (!existingKeys.has(key)) {
              types.push(node);
              existingKeys.add(key);
            }
          }
        }
      }
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
   * Generate all table related files
   */
  private _updateTableFiles(): void {
    // Collect all tables with their global indices
    let globalTableIndex = 1;
    const sheetTables: Map<string, { table: import('./table').Table; globalIndex: number }[]> = new Map();

    for (const def of this._sheetDefs) {
      const worksheet = this._sheets.get(def.name);
      if (!worksheet) continue;

      const tables = worksheet.tables;
      if (tables.length === 0) continue;

      const tableInfos: { table: import('./table').Table; globalIndex: number }[] = [];
      for (const table of tables) {
        tableInfos.push({ table, globalIndex: globalTableIndex });
        globalTableIndex++;
      }
      sheetTables.set(def.name, tableInfos);
    }

    // Generate table files
    for (const [, tableInfos] of sheetTables) {
      for (const { table, globalIndex } of tableInfos) {
        const tablePath = `xl/tables/table${globalIndex}.xml`;
        writeZipText(this._files, tablePath, table.toXml());
      }
    }

    // Generate worksheet relationships for tables
    for (const [sheetName, tableInfos] of sheetTables) {
      const def = this._sheetDefs.find((s) => s.name === sheetName);
      if (!def) continue;

      const rel = this._relationships.find((r) => r.id === def.rId);
      if (!rel) continue;

      // Extract sheet file name from target path
      const sheetFileName = rel.target.split('/').pop();
      const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;

      // Check if there are already pivot table relationships for this sheet
      const existingRelsXml = readZipText(this._files, sheetRelsPath);
      let nextRelId = 1;
      const relNodes: XmlNode[] = [];
      const reservedRelIds = new Set<string>();

      if (existingRelsXml) {
        // Parse existing rels and find max rId
        const parsed = parseXml(existingRelsXml);
        const relsElement = findElement(parsed, 'Relationships');
        if (relsElement) {
          const existingRelNodes = getChildren(relsElement, 'Relationships');
          for (const relNode of existingRelNodes) {
            if ('Relationship' in relNode) {
              relNodes.push(relNode);
              const id = getAttr(relNode, 'Id');
              if (id) {
                reservedRelIds.add(id);
                const idNum = parseInt(id.replace('rId', ''), 10);
                if (idNum >= nextRelId) {
                  nextRelId = idNum + 1;
                }
              }
            }
          }
        }
      }

      const allocateRelId = (): string => {
        while (reservedRelIds.has(`rId${nextRelId}`)) {
          nextRelId++;
        }
        const id = `rId${nextRelId}`;
        nextRelId++;
        reservedRelIds.add(id);
        return id;
      };

      // Add table relationships
      const tableRelIds: string[] = [];
      for (const { globalIndex } of tableInfos) {
        const target = `../tables/table${globalIndex}.xml`;
        const existing = relNodes.some(
          (node) =>
            getAttr(node, 'Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table' &&
            getAttr(node, 'Target') === target,
        );
        if (existing) {
          const existingRel = relNodes.find(
            (node) =>
              getAttr(node, 'Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table' &&
              getAttr(node, 'Target') === target,
          );
          const existingId = existingRel ? getAttr(existingRel, 'Id') : undefined;
          tableRelIds.push(existingId ?? allocateRelId());
          continue;
        }
        const id = allocateRelId();
        tableRelIds.push(id);
        relNodes.push(
          createElement(
            'Relationship',
            {
              Id: id,
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
              Target: target,
            },
            [],
          ),
        );
      }

      const worksheet = this._sheets.get(sheetName);
      if (worksheet) {
        worksheet.setTableRelIds(tableRelIds);
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

  /**
   * Generate pivot cache/table parts and worksheet relationships.
   */
  private _updatePivotFiles(): void {
    if (this._pivotTables.length === 0) {
      return;
    }

    for (const pivot of this._pivotTables) {
      const pivotCachePath = `xl/pivotCache/pivotCacheDefinition${pivot.cachePartIndex}.xml`;
      writeZipText(this._files, pivotCachePath, pivot.toPivotCacheDefinitionXml());

      const pivotCacheRecordsPath = `xl/pivotCache/pivotCacheRecords${pivot.cachePartIndex}.xml`;
      writeZipText(this._files, pivotCacheRecordsPath, pivot.toPivotCacheRecordsXml());

      const pivotCacheRelsPath = `xl/pivotCache/_rels/pivotCacheDefinition${pivot.cachePartIndex}.xml.rels`;
      writeZipText(this._files, pivotCacheRelsPath, pivot.toPivotCacheDefinitionRelsXml());

      const pivotTablePath = `xl/pivotTables/pivotTable${pivot.pivotId}.xml`;
      writeZipText(this._files, pivotTablePath, pivot.toPivotTableDefinitionXml());
    }

    const pivotsBySheet = new Map<string, PivotTable[]>();
    for (const pivot of this._pivotTables) {
      const existing = pivotsBySheet.get(pivot.targetSheetName) ?? [];
      existing.push(pivot);
      pivotsBySheet.set(pivot.targetSheetName, existing);
    }

    for (const [sheetName, pivots] of pivotsBySheet) {
      const def = this._sheetDefs.find((s) => s.name === sheetName);
      if (!def) continue;

      const rel = this._relationships.find((r) => r.id === def.rId);
      if (!rel) continue;

      const sheetFileName = rel.target.split('/').pop();
      if (!sheetFileName) continue;

      const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;
      const existingRelsXml = readZipText(this._files, sheetRelsPath);

      let nextRelId = 1;
      const relNodes: XmlNode[] = [];
      const reservedRelIds = new Set<string>();

      if (existingRelsXml) {
        const parsed = parseXml(existingRelsXml);
        const relsElement = findElement(parsed, 'Relationships');
        if (relsElement) {
          for (const relNode of getChildren(relsElement, 'Relationships')) {
            if ('Relationship' in relNode) {
              relNodes.push(relNode);
              const id = getAttr(relNode, 'Id');
              if (id) {
                reservedRelIds.add(id);
                const idNum = parseInt(id.replace('rId', ''), 10);
                if (idNum >= nextRelId) {
                  nextRelId = idNum + 1;
                }
              }
            }
          }
        }
      }

      const allocateRelId = (): string => {
        while (reservedRelIds.has(`rId${nextRelId}`)) {
          nextRelId++;
        }
        const id = `rId${nextRelId}`;
        nextRelId++;
        reservedRelIds.add(id);
        return id;
      };

      const pivotRelIds: string[] = [];
      for (const pivot of pivots) {
        const target = `../pivotTables/pivotTable${pivot.pivotId}.xml`;
        const existing = relNodes.find(
          (node) =>
            getAttr(node, 'Type') ===
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable' &&
            getAttr(node, 'Target') === target,
        );

        if (existing) {
          const existingId = getAttr(existing, 'Id');
          pivotRelIds.push(existingId ?? allocateRelId());
          continue;
        }

        const id = allocateRelId();
        pivotRelIds.push(id);
        relNodes.push(
          createElement(
            'Relationship',
            {
              Id: id,
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
              Target: target,
            },
            [],
          ),
        );
      }

      const worksheet = this._sheets.get(sheetName);
      if (worksheet) {
        worksheet.setPivotTableRelIds(pivotRelIds);
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
