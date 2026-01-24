import { readFile, writeFile } from 'fs/promises';
import type { SheetDefinition, Relationship } from './types';
import { Worksheet } from './worksheet';
import { SharedStrings } from './shared-strings';
import { Styles } from './styles';
import { readZip, writeZip, readZipText, writeZipText, ZipFiles } from './utils/zip';
import {
  parseXml,
  findElement,
  getChildren,
  getAttr,
  XmlNode,
  stringifyXml,
  createElement,
} from './utils/xml';

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

    // Add default sheet
    workbook.addSheet('Sheet1');

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

    // Update styles if modified
    if (this._styles.dirty) {
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
  }

  private _updateWorkbookXml(): void {
    const sheetNodes: XmlNode[] = this._sheetDefs.map((def) =>
      createElement('sheet', { name: def.name, sheetId: String(def.sheetId), 'r:id': def.rId }, [])
    );

    const sheetsNode = createElement('sheets', {}, sheetNodes);

    const workbookNode = createElement(
      'workbook',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      },
      [sheetsNode]
    );

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([workbookNode])}`;
    writeZipText(this._files, 'xl/workbook.xml', xml);
  }

  private _updateRelationshipsXml(): void {
    const relNodes: XmlNode[] = this._relationships.map((rel) =>
      createElement('Relationship', { Id: rel.id, Type: rel.type, Target: rel.target }, [])
    );

    // Add shared strings relationship if needed
    if (this._sharedStrings.count > 0) {
      const hasSharedStrings = this._relationships.some(
        (r) => r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
      );
      if (!hasSharedStrings) {
        const rId = `rId${this._relationships.length + 1}`;
        relNodes.push(
          createElement(
            'Relationship',
            {
              Id: rId,
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
              Target: 'sharedStrings.xml',
            },
            []
          )
        );
      }
    }

    // Add styles relationship if needed
    const hasStyles = this._relationships.some(
      (r) => r.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
    );
    if (!hasStyles) {
      const rId = `rId${this._relationships.length + 2}`;
      relNodes.push(
        createElement(
          'Relationship',
          {
            Id: rId,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            Target: 'styles.xml',
          },
          []
        )
      );
    }

    const relsNode = createElement(
      'Relationships',
      { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
      relNodes
    );

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([relsNode])}`;
    writeZipText(this._files, 'xl/_rels/workbook.xml.rels', xml);
  }

  private _updateContentTypes(): void {
    const types: XmlNode[] = [
      createElement('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }, []),
      createElement('Default', { Extension: 'xml', ContentType: 'application/xml' }, []),
      createElement('Override', { PartName: '/xl/workbook.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }, []),
      createElement('Override', { PartName: '/xl/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' }, []),
    ];

    // Add shared strings if present
    if (this._sharedStrings.count > 0) {
      types.push(
        createElement('Override', { PartName: '/xl/sharedStrings.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' }, [])
      );
    }

    // Add worksheets
    for (const def of this._sheetDefs) {
      const rel = this._relationships.find((r) => r.id === def.rId);
      if (rel) {
        types.push(
          createElement('Override', { PartName: `/xl/${rel.target}`, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' }, [])
        );
      }
    }

    const typesNode = createElement('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' }, types);

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
            []
          ),
        ]
      );
      writeZipText(this._files, '_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([rootRels])}`);
    }
  }
}
