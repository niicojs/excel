import type { CellData, RangeAddress } from './types';
import type { Workbook } from './workbook';
import { Cell, parseCellRef } from './cell';
import { Range } from './range';
import { parseRange, toAddress, parseAddress } from './utils/address';
import {
  parseXml,
  findElement,
  getChildren,
  getAttr,
  XmlNode,
  stringifyXml,
  createElement,
  createText,
} from './utils/xml';

/**
 * Represents a worksheet in a workbook
 */
export class Worksheet {
  private _name: string;
  private _workbook: Workbook;
  private _cells: Map<string, Cell> = new Map();
  private _xmlNodes: XmlNode[] | null = null;
  private _dirty = false;
  private _mergedCells: Set<string> = new Set();
  private _sheetData: XmlNode[] = [];

  constructor(workbook: Workbook, name: string) {
    this._workbook = workbook;
    this._name = name;
  }

  /**
   * Get the workbook this sheet belongs to
   */
  get workbook(): Workbook {
    return this._workbook;
  }

  /**
   * Get the sheet name
   */
  get name(): string {
    return this._name;
  }

  /**
   * Set the sheet name
   */
  set name(value: string) {
    this._name = value;
    this._dirty = true;
  }

  /**
   * Parse worksheet XML content
   */
  parse(xml: string): void {
    this._xmlNodes = parseXml(xml);
    const worksheet = findElement(this._xmlNodes, 'worksheet');
    if (!worksheet) return;

    const worksheetChildren = getChildren(worksheet, 'worksheet');

    // Parse sheet data (cells)
    const sheetData = findElement(worksheetChildren, 'sheetData');
    if (sheetData) {
      this._sheetData = getChildren(sheetData, 'sheetData');
      this._parseSheetData(this._sheetData);
    }

    // Parse merged cells
    const mergeCells = findElement(worksheetChildren, 'mergeCells');
    if (mergeCells) {
      const mergeChildren = getChildren(mergeCells, 'mergeCells');
      for (const mergeCell of mergeChildren) {
        if ('mergeCell' in mergeCell) {
          const ref = getAttr(mergeCell, 'ref');
          if (ref) {
            this._mergedCells.add(ref);
          }
        }
      }
    }
  }

  /**
   * Parse the sheetData element to extract cells
   */
  private _parseSheetData(rows: XmlNode[]): void {
    for (const rowNode of rows) {
      if (!('row' in rowNode)) continue;

      const rowChildren = getChildren(rowNode, 'row');
      for (const cellNode of rowChildren) {
        if (!('c' in cellNode)) continue;

        const ref = getAttr(cellNode, 'r');
        if (!ref) continue;

        const { row, col } = parseAddress(ref);
        const cellData = this._parseCellNode(cellNode);
        const cell = new Cell(this, row, col, cellData);
        this._cells.set(ref, cell);
      }
    }
  }

  /**
   * Parse a cell XML node to CellData
   */
  private _parseCellNode(node: XmlNode): CellData {
    const data: CellData = {};

    // Type attribute
    const t = getAttr(node, 't');
    if (t) {
      data.t = t as CellData['t'];
    }

    // Style attribute
    const s = getAttr(node, 's');
    if (s) {
      data.s = parseInt(s, 10);
    }

    const children = getChildren(node, 'c');

    // Value element
    const vNode = findElement(children, 'v');
    if (vNode) {
      const vChildren = getChildren(vNode, 'v');
      for (const child of vChildren) {
        if ('#text' in child) {
          const text = child['#text'] as string;
          // Parse based on type
          if (data.t === 's') {
            data.v = parseInt(text, 10); // Shared string index
          } else if (data.t === 'b') {
            data.v = text === '1' ? 1 : 0;
          } else if (data.t === 'e' || data.t === 'str') {
            data.v = text;
          } else {
            // Number or default
            data.v = parseFloat(text);
          }
          break;
        }
      }
    }

    // Formula element
    const fNode = findElement(children, 'f');
    if (fNode) {
      const fChildren = getChildren(fNode, 'f');
      for (const child of fChildren) {
        if ('#text' in child) {
          data.f = child['#text'] as string;
          break;
        }
      }

      // Check for shared formula
      const si = getAttr(fNode, 'si');
      if (si) {
        data.si = parseInt(si, 10);
      }

      // Check for array formula range
      const ref = getAttr(fNode, 'ref');
      if (ref) {
        data.F = ref;
      }
    }

    // Inline string (is element)
    const isNode = findElement(children, 'is');
    if (isNode) {
      data.t = 'str';
      const isChildren = getChildren(isNode, 'is');
      const tNode = findElement(isChildren, 't');
      if (tNode) {
        const tChildren = getChildren(tNode, 't');
        for (const child of tChildren) {
          if ('#text' in child) {
            data.v = child['#text'] as string;
            break;
          }
        }
      }
    }

    return data;
  }

  /**
   * Get a cell by address or row/col
   */
  cell(rowOrAddress: number | string, col?: number): Cell {
    const { row, col: c } = parseCellRef(rowOrAddress, col);
    const address = toAddress(row, c);

    let cell = this._cells.get(address);
    if (!cell) {
      cell = new Cell(this, row, c);
      this._cells.set(address, cell);
    }

    return cell;
  }

  /**
   * Get a range of cells
   */
  range(rangeStr: string): Range;
  range(startRow: number, startCol: number, endRow: number, endCol: number): Range;
  range(startRowOrRange: number | string, startCol?: number, endRow?: number, endCol?: number): Range {
    let rangeAddr: RangeAddress;

    if (typeof startRowOrRange === 'string') {
      rangeAddr = parseRange(startRowOrRange);
    } else {
      if (startCol === undefined || endRow === undefined || endCol === undefined) {
        throw new Error('All range parameters must be provided');
      }
      rangeAddr = {
        start: { row: startRowOrRange, col: startCol },
        end: { row: endRow, col: endCol },
      };
    }

    return new Range(this, rangeAddr);
  }

  /**
   * Merge cells in the given range
   */
  mergeCells(rangeOrStart: string, end?: string): void {
    let rangeStr: string;
    if (end) {
      rangeStr = `${rangeOrStart}:${end}`;
    } else {
      rangeStr = rangeOrStart;
    }
    this._mergedCells.add(rangeStr);
    this._dirty = true;
  }

  /**
   * Unmerge cells in the given range
   */
  unmergeCells(rangeStr: string): void {
    this._mergedCells.delete(rangeStr);
    this._dirty = true;
  }

  /**
   * Get all merged cell ranges
   */
  get mergedCells(): string[] {
    return Array.from(this._mergedCells);
  }

  /**
   * Check if the worksheet has been modified
   */
  get dirty(): boolean {
    if (this._dirty) return true;
    for (const cell of this._cells.values()) {
      if (cell.dirty) return true;
    }
    return false;
  }

  /**
   * Get all cells in the worksheet
   */
  get cells(): Map<string, Cell> {
    return this._cells;
  }

  /**
   * Generate XML for this worksheet
   */
  toXml(): string {
    // Build sheetData from cells
    const rowMap = new Map<number, Cell[]>();
    for (const cell of this._cells.values()) {
      const row = cell.row;
      if (!rowMap.has(row)) {
        rowMap.set(row, []);
      }
      rowMap.get(row)!.push(cell);
    }

    // Sort rows and cells
    const sortedRows = Array.from(rowMap.entries()).sort((a, b) => a[0] - b[0]);

    const rowNodes: XmlNode[] = [];
    for (const [rowIdx, cells] of sortedRows) {
      cells.sort((a, b) => a.col - b.col);

      const cellNodes: XmlNode[] = [];
      for (const cell of cells) {
        const cellNode = this._buildCellNode(cell);
        cellNodes.push(cellNode);
      }

      const rowNode = createElement('row', { r: String(rowIdx + 1) }, cellNodes);
      rowNodes.push(rowNode);
    }

    const sheetDataNode = createElement('sheetData', {}, rowNodes);

    // Build worksheet structure
    const worksheetChildren: XmlNode[] = [sheetDataNode];

    // Add merged cells if any
    if (this._mergedCells.size > 0) {
      const mergeCellNodes: XmlNode[] = [];
      for (const ref of this._mergedCells) {
        mergeCellNodes.push(createElement('mergeCell', { ref }, []));
      }
      const mergeCellsNode = createElement('mergeCells', { count: String(this._mergedCells.size) }, mergeCellNodes);
      worksheetChildren.push(mergeCellsNode);
    }

    const worksheetNode = createElement(
      'worksheet',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      },
      worksheetChildren,
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([worksheetNode])}`;
  }

  /**
   * Build a cell XML node from a Cell object
   */
  private _buildCellNode(cell: Cell): XmlNode {
    const data = cell.data;
    const attrs: Record<string, string> = { r: cell.address };

    if (data.t && data.t !== 'n') {
      attrs.t = data.t;
    }
    if (data.s !== undefined) {
      attrs.s = String(data.s);
    }

    const children: XmlNode[] = [];

    // Formula
    if (data.f) {
      const fAttrs: Record<string, string> = {};
      if (data.F) fAttrs.ref = data.F;
      if (data.si !== undefined) fAttrs.si = String(data.si);
      children.push(createElement('f', fAttrs, [createText(data.f)]));
    }

    // Value
    if (data.v !== undefined) {
      children.push(createElement('v', {}, [createText(String(data.v))]));
    }

    return createElement('c', attrs, children);
  }
}
