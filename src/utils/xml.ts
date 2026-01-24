import { XMLParser, XMLBuilder } from 'fast-xml-parser';

// Parser options that preserve structure and attributes
const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  preserveOrder: true,
  commentPropName: '#comment',
  cdataPropName: '#cdata',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
};

// Builder options matching parser for round-trip compatibility
const builderOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  preserveOrder: true,
  commentPropName: '#comment',
  cdataPropName: '#cdata',
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

const parser = new XMLParser(parserOptions);
const builder = new XMLBuilder(builderOptions);

/**
 * Parses an XML string into a JavaScript object
 * Preserves element order and attributes for round-trip compatibility
 */
export const parseXml = (xml: string): XmlNode[] => {
  return parser.parse(xml);
};

/**
 * Converts a JavaScript object back to an XML string
 */
export const stringifyXml = (obj: XmlNode[]): string => {
  return builder.build(obj);
};

/**
 * XML node type from fast-xml-parser with preserveOrder
 * Each node is an object with a single key (the tag name)
 * containing an array of child nodes, plus optional :@
 * for attributes
 */
export interface XmlNode {
  [tagName: string]: XmlNode[] | string | Record<string, string> | undefined;
}

/**
 * Finds the first element with the given tag name in the XML tree
 */
export const findElement = (nodes: XmlNode[], tagName: string): XmlNode | undefined => {
  for (const node of nodes) {
    if (tagName in node) {
      return node;
    }
  }
  return undefined;
};

/**
 * Finds all elements with the given tag name (immediate children only)
 */
export const findElements = (nodes: XmlNode[], tagName: string): XmlNode[] => {
  return nodes.filter((node) => tagName in node);
};

/**
 * Gets the children of an element
 */
export const getChildren = (node: XmlNode, tagName: string): XmlNode[] => {
  const children = node[tagName];
  if (Array.isArray(children)) {
    return children;
  }
  return [];
};

/**
 * Gets an attribute value from a node
 */
export const getAttr = (node: XmlNode, name: string): string | undefined => {
  const attrs = node[':@'] as Record<string, string> | undefined;
  return attrs?.[`@_${name}`];
};

/**
 * Sets an attribute value on a node
 */
export const setAttr = (node: XmlNode, name: string, value: string): void => {
  if (!node[':@']) {
    node[':@'] = {};
  }
  (node[':@'] as Record<string, string>)[`@_${name}`] = value;
};

/**
 * Gets the text content of a node
 */
export const getText = (node: XmlNode, tagName: string): string | undefined => {
  const children = getChildren(node, tagName);
  for (const child of children) {
    if ('#text' in child) {
      return child['#text'] as string;
    }
  }
  return undefined;
};

/**
 * Creates a new XML element
 */
export const createElement = (
  tagName: string,
  attrs?: Record<string, string>,
  children?: XmlNode[]
): XmlNode => {
  const node: XmlNode = {
    [tagName]: children || [],
  };
  if (attrs && Object.keys(attrs).length > 0) {
    const attrObj: Record<string, string> = {};
    for (const [key, value] of Object.entries(attrs)) {
      attrObj[`@_${key}`] = value;
    }
    node[':@'] = attrObj;
  }
  return node;
};

/**
 * Creates a text node
 */
export const createText = (text: string): XmlNode => {
  return { '#text': text } as unknown as XmlNode;
};

/**
 * Adds XML declaration to the start of an XML string
 */
export const addXmlDeclaration = (xml: string): string => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
};
