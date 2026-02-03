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
 * Manages the shared strings table from xl/sharedStrings.xml
 * Excel stores strings in a shared table to reduce file size
 */
export class SharedStrings {
  private entries: SharedStringEntry[] = [];
  private stringToIndex: Map<string, number> = new Map();
  private _dirty = false;
  private _totalCount = 0;

  /**
   * Parse shared strings from XML content
   */
  static parse(xml: string): SharedStrings {
    const ss = new SharedStrings();
    const parsed = parseXml(xml);
    const sst = findElement(parsed, 'sst');
    if (!sst) return ss;

    const countAttr = getAttr(sst, 'count');
    if (countAttr) {
      const total = parseInt(countAttr, 10);
      if (Number.isFinite(total) && total >= 0) {
        ss._totalCount = total;
      }
    }

    const children = getChildren(sst, 'sst');
    for (const child of children) {
      if ('si' in child) {
        const siChildren = getChildren(child, 'si');
        const text = ss.extractText(siChildren);
        ss.entries.push({ text, node: child });
        ss.stringToIndex.set(text, ss.entries.length - 1);
      }
    }

    if (ss._totalCount === 0 && ss.entries.length > 0) {
      ss._totalCount = ss.entries.length;
    }

    return ss;
  }

  /**
   * Extract text from a string item (si element)
   * Handles both simple <t> elements and rich text <r> elements
   */
  private extractText(nodes: XmlNode[]): string {
    let text = '';
    for (const node of nodes) {
      if ('t' in node) {
        // Simple text: <t>value</t>
        const tChildren = getChildren(node, 't');
        for (const child of tChildren) {
          if ('#text' in child) {
            text += child['#text'] as string;
          }
        }
      } else if ('r' in node) {
        // Rich text: <r><t>value</t></r>
        const rChildren = getChildren(node, 'r');
        for (const rChild of rChildren) {
          if ('t' in rChild) {
            const tChildren = getChildren(rChild, 't');
            for (const child of tChildren) {
              if ('#text' in child) {
                text += child['#text'] as string;
              }
            }
          }
        }
      }
    }
    return text;
  }

  /**
   * Get a string by index
   */
  getString(index: number): string | undefined {
    return this.entries[index]?.text;
  }

  /**
   * Add a string and return its index
   * If the string already exists, returns the existing index
   */
  addString(str: string): number {
    const existing = this.stringToIndex.get(str);
    if (existing !== undefined) {
      this._totalCount++;
      this._dirty = true;
      return existing;
    }
    const index = this.entries.length;
    const tElement = createElement('t', str.startsWith(' ') || str.endsWith(' ') ? { 'xml:space': 'preserve' } : {}, [
      createText(str),
    ]);
    const siElement = createElement('si', {}, [tElement]);
    this.entries.push({ text: str, node: siElement });
    this.stringToIndex.set(str, index);
    this._totalCount++;
    this._dirty = true;
    return index;
  }

  /**
   * Check if the shared strings table has been modified
   */
  get dirty(): boolean {
    return this._dirty;
  }

  /**
   * Get the count of strings
   */
  get count(): number {
    return this.entries.length;
  }

  /**
   * Get total usage count of shared strings
   */
  get totalCount(): number {
    return Math.max(this._totalCount, this.entries.length);
  }

  /**
   * Generate XML for the shared strings table
   */
  toXml(): string {
    const siElements: XmlNode[] = [];
    for (const entry of this.entries) {
      if (entry.node) {
        siElements.push(entry.node);
      } else {
        const str = entry.text;
        const tElement = createElement(
          't',
          str.startsWith(' ') || str.endsWith(' ') ? { 'xml:space': 'preserve' } : {},
          [createText(str)],
        );
        const siElement = createElement('si', {}, [tElement]);
        siElements.push(siElement);
      }
    }

    const totalCount = Math.max(this._totalCount, this.entries.length);
    const sst = createElement(
      'sst',
      {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        count: String(totalCount),
        uniqueCount: String(this.entries.length),
      },
      siElements,
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([sst])}`;
  }
}

interface SharedStringEntry {
  text: string;
  node?: XmlNode;
}
