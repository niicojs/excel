import { describe, it, expect } from 'vitest';
import {
  colToLetter,
  letterToCol,
  parseAddress,
  toAddress,
  parseRange,
  toRange,
  normalizeRange,
  isInRange,
} from '../../src/utils/address';

describe('colToLetter', () => {
  it('converts single letter columns', () => {
    expect(colToLetter(0)).toBe('A');
    expect(colToLetter(1)).toBe('B');
    expect(colToLetter(25)).toBe('Z');
  });

  it('converts double letter columns', () => {
    expect(colToLetter(26)).toBe('AA');
    expect(colToLetter(27)).toBe('AB');
    expect(colToLetter(51)).toBe('AZ');
    expect(colToLetter(52)).toBe('BA');
  });

  it('converts triple letter columns', () => {
    expect(colToLetter(702)).toBe('AAA');
  });
});

describe('letterToCol', () => {
  it('converts single letter columns', () => {
    expect(letterToCol('A')).toBe(0);
    expect(letterToCol('B')).toBe(1);
    expect(letterToCol('Z')).toBe(25);
  });

  it('converts double letter columns', () => {
    expect(letterToCol('AA')).toBe(26);
    expect(letterToCol('AB')).toBe(27);
    expect(letterToCol('AZ')).toBe(51);
    expect(letterToCol('BA')).toBe(52);
  });

  it('handles lowercase', () => {
    expect(letterToCol('a')).toBe(0);
    expect(letterToCol('aa')).toBe(26);
  });
});

describe('parseAddress', () => {
  it('parses simple addresses', () => {
    expect(parseAddress('A1')).toEqual({ row: 0, col: 0 });
    expect(parseAddress('B2')).toEqual({ row: 1, col: 1 });
    expect(parseAddress('Z100')).toEqual({ row: 99, col: 25 });
  });

  it('parses addresses with absolute references', () => {
    expect(parseAddress('$A$1')).toEqual({ row: 0, col: 0 });
    expect(parseAddress('$B2')).toEqual({ row: 1, col: 1 });
    expect(parseAddress('B$2')).toEqual({ row: 1, col: 1 });
  });

  it('handles multi-letter columns', () => {
    expect(parseAddress('AA1')).toEqual({ row: 0, col: 26 });
    expect(parseAddress('AAA1')).toEqual({ row: 0, col: 702 });
  });

  it('throws on invalid addresses', () => {
    expect(() => parseAddress('')).toThrow();
    expect(() => parseAddress('1')).toThrow();
    expect(() => parseAddress('A')).toThrow();
    expect(() => parseAddress('1A')).toThrow();
  });
});

describe('toAddress', () => {
  it('converts row/col to address', () => {
    expect(toAddress(0, 0)).toBe('A1');
    expect(toAddress(1, 1)).toBe('B2');
    expect(toAddress(99, 25)).toBe('Z100');
  });

  it('handles multi-letter columns', () => {
    expect(toAddress(0, 26)).toBe('AA1');
    expect(toAddress(0, 702)).toBe('AAA1');
  });
});

describe('parseRange', () => {
  it('parses simple ranges', () => {
    expect(parseRange('A1:B2')).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 1, col: 1 },
    });
  });

  it('parses single cell as range', () => {
    expect(parseRange('A1')).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 0, col: 0 },
    });
  });

  it('handles absolute references', () => {
    expect(parseRange('$A$1:$B$2')).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 1, col: 1 },
    });
  });
});

describe('toRange', () => {
  it('converts range to string', () => {
    expect(toRange({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } })).toBe('A1:B2');
  });

  it('returns single cell for same start/end', () => {
    expect(toRange({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } })).toBe('A1');
  });
});

describe('normalizeRange', () => {
  it('normalizes already-normalized range', () => {
    const range = { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } };
    expect(normalizeRange(range)).toEqual(range);
  });

  it('normalizes reversed range', () => {
    expect(normalizeRange({ start: { row: 1, col: 1 }, end: { row: 0, col: 0 } })).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 1, col: 1 },
    });
  });
});

describe('isInRange', () => {
  const range = { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } };

  it('returns true for cells in range', () => {
    expect(isInRange({ row: 0, col: 0 }, range)).toBe(true);
    expect(isInRange({ row: 1, col: 1 }, range)).toBe(true);
    expect(isInRange({ row: 2, col: 2 }, range)).toBe(true);
  });

  it('returns false for cells outside range', () => {
    expect(isInRange({ row: 3, col: 0 }, range)).toBe(false);
    expect(isInRange({ row: 0, col: 3 }, range)).toBe(false);
  });
});
