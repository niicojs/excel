import type { CellAddress, RangeAddress } from '../types';

/**
 * Converts a column index (0-based) to Excel column letters (A, B, ..., Z, AA, AB, ...)
 * @param col - 0-based column index
 * @returns Column letter(s)
 */
export const colToLetter = (col: number): string => {
  let result = '';
  let n = col;
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
};

/**
 * Converts Excel column letters to a 0-based column index
 * @param letters - Column letter(s) like 'A', 'B', 'AA'
 * @returns 0-based column index
 */
export const letterToCol = (letters: string): number => {
  const upper = letters.toUpperCase();
  let col = 0;
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  return col - 1;
};

/**
 * Parses an Excel cell address (e.g., 'A1', '$B$2') to row/col indices
 * @param address - Cell address string
 * @returns CellAddress with 0-based row and col
 */
export const parseAddress = (address: string): CellAddress => {
  // Remove $ signs for absolute references
  const clean = address.replace(/\$/g, '');
  const match = clean.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  const col = letterToCol(match[1].toUpperCase());
  const row = parseInt(match[2], 10) - 1; // Convert to 0-based
  return { row, col };
};

/**
 * Converts row/col indices to an Excel cell address
 * @param row - 0-based row index
 * @param col - 0-based column index
 * @returns Cell address string like 'A1'
 */
export const toAddress = (row: number, col: number): string => {
  return `${colToLetter(col)}${row + 1}`;
};

/**
 * Parses an Excel range (e.g., 'A1:B10') to start/end addresses
 * @param range - Range string
 * @returns RangeAddress with start and end
 */
export const parseRange = (range: string): RangeAddress => {
  const parts = range.split(':');
  if (parts.length === 1) {
    // Single cell range
    const addr = parseAddress(parts[0]);
    return { start: addr, end: addr };
  }
  if (parts.length !== 2) {
    throw new Error(`Invalid range: ${range}`);
  }
  return {
    start: parseAddress(parts[0]),
    end: parseAddress(parts[1]),
  };
};

/**
 * Converts a RangeAddress to a range string
 * @param range - RangeAddress object
 * @returns Range string like 'A1:B10'
 */
export const toRange = (range: RangeAddress): string => {
  const start = toAddress(range.start.row, range.start.col);
  const end = toAddress(range.end.row, range.end.col);
  if (start === end) {
    return start;
  }
  return `${start}:${end}`;
};

/**
 * Normalizes a range so start is always top-left and end is bottom-right
 */
export const normalizeRange = (range: RangeAddress): RangeAddress => {
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
};

/**
 * Checks if an address is within a range
 */
export const isInRange = (addr: CellAddress, range: RangeAddress): boolean => {
  const norm = normalizeRange(range);
  return (
    addr.row >= norm.start.row && addr.row <= norm.end.row && addr.col >= norm.start.col && addr.col <= norm.end.col
  );
};
