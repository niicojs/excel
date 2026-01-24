import type { CellStyle, BorderType } from './types';
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
 * Manages the styles (xl/styles.xml)
 */
export class Styles {
  private _numFmts: Map<number, string> = new Map();
  private _fonts: StyleFont[] = [];
  private _fills: StyleFill[] = [];
  private _borders: StyleBorder[] = [];
  private _cellXfs: CellXf[] = []; // Cell formats (combined style index)
  private _xmlNodes: XmlNode[] | null = null;
  private _dirty = false;

  // Cache for style deduplication
  private _styleCache: Map<string, number> = new Map();

  /**
   * Parse styles from XML content
   */
  static parse(xml: string): Styles {
    const styles = new Styles();
    styles._xmlNodes = parseXml(xml);

    const styleSheet = findElement(styles._xmlNodes, 'styleSheet');
    if (!styleSheet) return styles;

    const children = getChildren(styleSheet, 'styleSheet');

    // Parse number formats
    const numFmts = findElement(children, 'numFmts');
    if (numFmts) {
      for (const child of getChildren(numFmts, 'numFmts')) {
        if ('numFmt' in child) {
          const id = parseInt(getAttr(child, 'numFmtId') || '0', 10);
          const code = getAttr(child, 'formatCode') || '';
          styles._numFmts.set(id, code);
        }
      }
    }

    // Parse fonts
    const fonts = findElement(children, 'fonts');
    if (fonts) {
      for (const child of getChildren(fonts, 'fonts')) {
        if ('font' in child) {
          styles._fonts.push(styles._parseFont(child));
        }
      }
    }

    // Parse fills
    const fills = findElement(children, 'fills');
    if (fills) {
      for (const child of getChildren(fills, 'fills')) {
        if ('fill' in child) {
          styles._fills.push(styles._parseFill(child));
        }
      }
    }

    // Parse borders
    const borders = findElement(children, 'borders');
    if (borders) {
      for (const child of getChildren(borders, 'borders')) {
        if ('border' in child) {
          styles._borders.push(styles._parseBorder(child));
        }
      }
    }

    // Parse cellXfs (cell formats)
    const cellXfs = findElement(children, 'cellXfs');
    if (cellXfs) {
      for (const child of getChildren(cellXfs, 'cellXfs')) {
        if ('xf' in child) {
          styles._cellXfs.push(styles._parseCellXf(child));
        }
      }
    }

    return styles;
  }

  /**
   * Create an empty styles object with defaults
   */
  static createDefault(): Styles {
    const styles = new Styles();

    // Default font (Calibri 11)
    styles._fonts.push({
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      size: 11,
      name: 'Calibri',
      color: undefined,
    });

    // Default fills (none and gray125 pattern are required)
    styles._fills.push({ type: 'none' });
    styles._fills.push({ type: 'gray125' });

    // Default border (none)
    styles._borders.push({});

    // Default cell format
    styles._cellXfs.push({
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId: 0,
    });

    return styles;
  }

  private _parseFont(node: XmlNode): StyleFont {
    const font: StyleFont = {
      bold: false,
      italic: false,
      underline: false,
      strike: false,
    };

    const children = getChildren(node, 'font');
    for (const child of children) {
      if ('b' in child) font.bold = true;
      if ('i' in child) font.italic = true;
      if ('u' in child) font.underline = true;
      if ('strike' in child) font.strike = true;
      if ('sz' in child) font.size = parseFloat(getAttr(child, 'val') || '11');
      if ('name' in child) font.name = getAttr(child, 'val');
      if ('color' in child) {
        font.color = getAttr(child, 'rgb') || getAttr(child, 'theme');
      }
    }

    return font;
  }

  private _parseFill(node: XmlNode): StyleFill {
    const fill: StyleFill = { type: 'none' };
    const children = getChildren(node, 'fill');

    for (const child of children) {
      if ('patternFill' in child) {
        const pattern = getAttr(child, 'patternType');
        fill.type = pattern || 'none';

        const pfChildren = getChildren(child, 'patternFill');
        for (const pfChild of pfChildren) {
          if ('fgColor' in pfChild) {
            fill.fgColor = getAttr(pfChild, 'rgb') || getAttr(pfChild, 'theme');
          }
          if ('bgColor' in pfChild) {
            fill.bgColor = getAttr(pfChild, 'rgb') || getAttr(pfChild, 'theme');
          }
        }
      }
    }

    return fill;
  }

  private _parseBorder(node: XmlNode): StyleBorder {
    const border: StyleBorder = {};
    const children = getChildren(node, 'border');

    for (const child of children) {
      const style = getAttr(child, 'style') as BorderType | undefined;
      if ('left' in child && style) border.left = style;
      if ('right' in child && style) border.right = style;
      if ('top' in child && style) border.top = style;
      if ('bottom' in child && style) border.bottom = style;
    }

    return border;
  }

  private _parseCellXf(node: XmlNode): CellXf {
    return {
      fontId: parseInt(getAttr(node, 'fontId') || '0', 10),
      fillId: parseInt(getAttr(node, 'fillId') || '0', 10),
      borderId: parseInt(getAttr(node, 'borderId') || '0', 10),
      numFmtId: parseInt(getAttr(node, 'numFmtId') || '0', 10),
      alignment: this._parseAlignment(node),
    };
  }

  private _parseAlignment(node: XmlNode): AlignmentStyle | undefined {
    const children = getChildren(node, 'xf');
    const alignNode = findElement(children, 'alignment');
    if (!alignNode) return undefined;

    return {
      horizontal: getAttr(alignNode, 'horizontal') as AlignmentStyle['horizontal'],
      vertical: getAttr(alignNode, 'vertical') as AlignmentStyle['vertical'],
      wrapText: getAttr(alignNode, 'wrapText') === '1',
      textRotation: parseInt(getAttr(alignNode, 'textRotation') || '0', 10),
    };
  }

  /**
   * Get a style by index
   */
  getStyle(index: number): CellStyle {
    const xf = this._cellXfs[index];
    if (!xf) return {};

    const font = this._fonts[xf.fontId];
    const fill = this._fills[xf.fillId];
    const border = this._borders[xf.borderId];
    const numFmt = this._numFmts.get(xf.numFmtId);

    const style: CellStyle = {};

    if (font) {
      if (font.bold) style.bold = true;
      if (font.italic) style.italic = true;
      if (font.underline) style.underline = true;
      if (font.strike) style.strike = true;
      if (font.size) style.fontSize = font.size;
      if (font.name) style.fontName = font.name;
      if (font.color) style.fontColor = font.color;
    }

    if (fill && fill.fgColor) {
      style.fill = fill.fgColor;
    }

    if (border) {
      if (border.top || border.bottom || border.left || border.right) {
        style.border = {
          top: border.top,
          bottom: border.bottom,
          left: border.left,
          right: border.right,
        };
      }
    }

    if (numFmt) {
      style.numberFormat = numFmt;
    }

    if (xf.alignment) {
      style.alignment = {
        horizontal: xf.alignment.horizontal,
        vertical: xf.alignment.vertical,
        wrapText: xf.alignment.wrapText,
        textRotation: xf.alignment.textRotation,
      };
    }

    return style;
  }

  /**
   * Create a style and return its index
   * Uses caching to deduplicate identical styles
   */
  createStyle(style: CellStyle): number {
    const key = JSON.stringify(style);
    const cached = this._styleCache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    this._dirty = true;

    // Create or find font
    const fontId = this._findOrCreateFont(style);

    // Create or find fill
    const fillId = this._findOrCreateFill(style);

    // Create or find border
    const borderId = this._findOrCreateBorder(style);

    // Create or find number format
    const numFmtId = style.numberFormat ? this._findOrCreateNumFmt(style.numberFormat) : 0;

    // Create cell format
    const xf: CellXf = {
      fontId,
      fillId,
      borderId,
      numFmtId,
    };

    if (style.alignment) {
      xf.alignment = {
        horizontal: style.alignment.horizontal,
        vertical: style.alignment.vertical,
        wrapText: style.alignment.wrapText,
        textRotation: style.alignment.textRotation,
      };
    }

    const index = this._cellXfs.length;
    this._cellXfs.push(xf);
    this._styleCache.set(key, index);

    return index;
  }

  private _findOrCreateFont(style: CellStyle): number {
    const font: StyleFont = {
      bold: style.bold || false,
      italic: style.italic || false,
      underline: style.underline === true || style.underline === 'single' || style.underline === 'double',
      strike: style.strike || false,
      size: style.fontSize,
      name: style.fontName,
      color: style.fontColor,
    };

    // Try to find existing font
    for (let i = 0; i < this._fonts.length; i++) {
      const f = this._fonts[i];
      if (
        f.bold === font.bold &&
        f.italic === font.italic &&
        f.underline === font.underline &&
        f.strike === font.strike &&
        f.size === font.size &&
        f.name === font.name &&
        f.color === font.color
      ) {
        return i;
      }
    }

    // Create new font
    this._fonts.push(font);
    return this._fonts.length - 1;
  }

  private _findOrCreateFill(style: CellStyle): number {
    if (!style.fill) return 0;

    // Try to find existing fill
    for (let i = 0; i < this._fills.length; i++) {
      const f = this._fills[i];
      if (f.fgColor === style.fill) {
        return i;
      }
    }

    // Create new fill
    this._fills.push({
      type: 'solid',
      fgColor: style.fill,
    });
    return this._fills.length - 1;
  }

  private _findOrCreateBorder(style: CellStyle): number {
    if (!style.border) return 0;

    const border: StyleBorder = {
      top: style.border.top,
      bottom: style.border.bottom,
      left: style.border.left,
      right: style.border.right,
    };

    // Try to find existing border
    for (let i = 0; i < this._borders.length; i++) {
      const b = this._borders[i];
      if (
        b.top === border.top &&
        b.bottom === border.bottom &&
        b.left === border.left &&
        b.right === border.right
      ) {
        return i;
      }
    }

    // Create new border
    this._borders.push(border);
    return this._borders.length - 1;
  }

  private _findOrCreateNumFmt(format: string): number {
    // Check if already exists
    for (const [id, code] of this._numFmts) {
      if (code === format) return id;
    }

    // Create new (custom formats start at 164)
    const id = Math.max(164, ...Array.from(this._numFmts.keys())) + 1;
    this._numFmts.set(id, format);
    return id;
  }

  /**
   * Check if styles have been modified
   */
  get dirty(): boolean {
    return this._dirty;
  }

  /**
   * Generate XML for styles
   */
  toXml(): string {
    const children: XmlNode[] = [];

    // Number formats
    if (this._numFmts.size > 0) {
      const numFmtNodes: XmlNode[] = [];
      for (const [id, code] of this._numFmts) {
        numFmtNodes.push(createElement('numFmt', { numFmtId: String(id), formatCode: code }, []));
      }
      children.push(createElement('numFmts', { count: String(numFmtNodes.length) }, numFmtNodes));
    }

    // Fonts
    const fontNodes: XmlNode[] = this._fonts.map((font) => this._buildFontNode(font));
    children.push(createElement('fonts', { count: String(fontNodes.length) }, fontNodes));

    // Fills
    const fillNodes: XmlNode[] = this._fills.map((fill) => this._buildFillNode(fill));
    children.push(createElement('fills', { count: String(fillNodes.length) }, fillNodes));

    // Borders
    const borderNodes: XmlNode[] = this._borders.map((border) => this._buildBorderNode(border));
    children.push(createElement('borders', { count: String(borderNodes.length) }, borderNodes));

    // Cell style xfs (required but we just add a default)
    children.push(
      createElement('cellStyleXfs', { count: '1' }, [createElement('xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' }, [])])
    );

    // Cell xfs
    const xfNodes: XmlNode[] = this._cellXfs.map((xf) => this._buildXfNode(xf));
    children.push(createElement('cellXfs', { count: String(xfNodes.length) }, xfNodes));

    // Cell styles (required)
    children.push(
      createElement('cellStyles', { count: '1' }, [
        createElement('cellStyle', { name: 'Normal', xfId: '0', builtinId: '0' }, []),
      ])
    );

    const styleSheet = createElement(
      'styleSheet',
      { xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' },
      children
    );

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([styleSheet])}`;
  }

  private _buildFontNode(font: StyleFont): XmlNode {
    const children: XmlNode[] = [];
    if (font.bold) children.push(createElement('b', {}, []));
    if (font.italic) children.push(createElement('i', {}, []));
    if (font.underline) children.push(createElement('u', {}, []));
    if (font.strike) children.push(createElement('strike', {}, []));
    if (font.size) children.push(createElement('sz', { val: String(font.size) }, []));
    if (font.color) children.push(createElement('color', { rgb: font.color }, []));
    if (font.name) children.push(createElement('name', { val: font.name }, []));
    return createElement('font', {}, children);
  }

  private _buildFillNode(fill: StyleFill): XmlNode {
    const patternChildren: XmlNode[] = [];
    if (fill.fgColor) {
      patternChildren.push(createElement('fgColor', { rgb: fill.fgColor }, []));
    }
    if (fill.bgColor) {
      patternChildren.push(createElement('bgColor', { rgb: fill.bgColor }, []));
    }
    const patternFill = createElement('patternFill', { patternType: fill.type || 'none' }, patternChildren);
    return createElement('fill', {}, [patternFill]);
  }

  private _buildBorderNode(border: StyleBorder): XmlNode {
    const children: XmlNode[] = [];
    if (border.left) children.push(createElement('left', { style: border.left }, []));
    if (border.right) children.push(createElement('right', { style: border.right }, []));
    if (border.top) children.push(createElement('top', { style: border.top }, []));
    if (border.bottom) children.push(createElement('bottom', { style: border.bottom }, []));
    // Add empty elements if not present (required by Excel)
    if (!border.left) children.push(createElement('left', {}, []));
    if (!border.right) children.push(createElement('right', {}, []));
    if (!border.top) children.push(createElement('top', {}, []));
    if (!border.bottom) children.push(createElement('bottom', {}, []));
    children.push(createElement('diagonal', {}, []));
    return createElement('border', {}, children);
  }

  private _buildXfNode(xf: CellXf): XmlNode {
    const attrs: Record<string, string> = {
      numFmtId: String(xf.numFmtId),
      fontId: String(xf.fontId),
      fillId: String(xf.fillId),
      borderId: String(xf.borderId),
    };

    if (xf.fontId > 0) attrs.applyFont = '1';
    if (xf.fillId > 0) attrs.applyFill = '1';
    if (xf.borderId > 0) attrs.applyBorder = '1';
    if (xf.numFmtId > 0) attrs.applyNumberFormat = '1';

    const children: XmlNode[] = [];
    if (xf.alignment) {
      const alignAttrs: Record<string, string> = {};
      if (xf.alignment.horizontal) alignAttrs.horizontal = xf.alignment.horizontal;
      if (xf.alignment.vertical) alignAttrs.vertical = xf.alignment.vertical;
      if (xf.alignment.wrapText) alignAttrs.wrapText = '1';
      if (xf.alignment.textRotation) alignAttrs.textRotation = String(xf.alignment.textRotation);
      children.push(createElement('alignment', alignAttrs, []));
      attrs.applyAlignment = '1';
    }

    return createElement('xf', attrs, children);
  }
}

// Internal types for style components
interface StyleFont {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  size?: number;
  name?: string;
  color?: string;
}

interface StyleFill {
  type: string;
  fgColor?: string;
  bgColor?: string;
}

interface StyleBorder {
  top?: BorderType;
  bottom?: BorderType;
  left?: BorderType;
  right?: BorderType;
}

interface CellXf {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: AlignmentStyle;
}

interface AlignmentStyle {
  horizontal?: 'left' | 'center' | 'right' | 'justify';
  vertical?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  textRotation?: number;
}
