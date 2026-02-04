import type { CellStyle } from '../types';

type NumberFormatInfo = {
  fractionDigits: number;
  percent: boolean;
  useGrouping: boolean;
  literalPrefix?: string;
  literalSuffix?: string;
};

type LocaleNumberInfo = {
  decimal: string;
  group: string;
};

const DEFAULT_LOCALE = 'fr-FR';

const formatPartsCache = new Map<string, Intl.NumberFormatPart[]>();

const getLocaleNumberInfo = (locale: string): LocaleNumberInfo => {
  const cacheKey = `num-info:${locale}`;
  const cached = formatPartsCache.get(cacheKey);
  if (cached) {
    const decimal = cached.find((p) => p.type === 'decimal')?.value ?? ',';
    const group = cached.find((p) => p.type === 'group')?.value ?? ' ';
    return { decimal, group };
  }

  const parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
  formatPartsCache.set(cacheKey, parts);
  const decimal = parts.find((p) => p.type === 'decimal')?.value ?? ',';
  const group = parts.find((p) => p.type === 'group')?.value ?? ' ';
  return { decimal, group };
};

const normalizeSpaces = (value: string): string => {
  return value.replace(/[\u00a0\u202f]/g, ' ');
};

const splitFormatSections = (format: string): string[] => {
  const sections = [] as string[];
  let current = '';
  let inQuote = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
      continue;
    }
    if (ch === ';' && !inQuote) {
      sections.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  sections.push(current);
  return sections;
};

const extractFormatLiterals = (format: string): { cleaned: string; prefix: string; suffix: string } => {
  let prefix = '';
  let suffix = '';
  let cleaned = '';
  let inQuote = false;
  let sawPlaceholder = false;

  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') {
      inQuote = !inQuote;
      continue;
    }
    if (inQuote) {
      if (!sawPlaceholder) {
        prefix += ch;
      } else {
        suffix += ch;
      }
      continue;
    }
    if (ch === '\\' && i + 1 < format.length) {
      const escaped = format[i + 1];
      if (!sawPlaceholder) {
        prefix += escaped;
      } else {
        suffix += escaped;
      }
      i++;
      continue;
    }
    if (ch === '_' || ch === '*') {
      if (i + 1 < format.length) {
        i++;
      }
      continue;
    }
    if (ch === '[') {
      const end = format.indexOf(']', i + 1);
      if (end !== -1) {
        const content = format.slice(i + 1, end);
        const currencyMatch = content.match(/[$€]/);
        if (currencyMatch) {
          if (!sawPlaceholder) {
            prefix += currencyMatch[0];
          } else {
            suffix += currencyMatch[0];
          }
        }
        i = end;
        continue;
      }
    }

    if (ch === '%') {
      if (!sawPlaceholder) {
        prefix += ch;
      } else {
        suffix += ch;
      }
      continue;
    }

    if (ch === '0' || ch === '#' || ch === '?' || ch === '.' || ch === ',') {
      sawPlaceholder = true;
      cleaned += ch;
      continue;
    }

    if (!sawPlaceholder) {
      prefix += ch;
    } else {
      suffix += ch;
    }
  }

  return { cleaned, prefix, suffix };
};

const parseNumberFormat = (format: string): NumberFormatInfo | null => {
  const trimmed = format.trim();
  if (!trimmed) return null;

  const { cleaned, prefix, suffix } = extractFormatLiterals(trimmed);
  const lower = cleaned.toLowerCase();
  if (!/[0#?]/.test(lower)) return null;

  const percent = /%/.test(trimmed);

  const section = lower;
  const lastDot = section.lastIndexOf('.');
  const lastComma = section.lastIndexOf(',');
  let decimalSeparator: '.' | ',' | null = null;

  if (lastDot >= 0 && lastComma >= 0) {
    decimalSeparator = lastDot > lastComma ? '.' : ',';
  } else if (lastDot >= 0 || lastComma >= 0) {
    const candidate = lastDot >= 0 ? '.' : ',';
    const index = lastDot >= 0 ? lastDot : lastComma;
    const fractionSection = section.slice(index + 1);
    const fractionDigitsCandidate = fractionSection.replace(/[^0#?]/g, '').length;
    if (fractionDigitsCandidate > 0) {
      if (fractionDigitsCandidate === 3) {
        decimalSeparator = null;
      } else {
        decimalSeparator = candidate;
      }
    }
  }

  let decimalIndex = decimalSeparator === '.' ? lastDot : decimalSeparator === ',' ? lastComma : -1;
  if (decimalIndex >= 0) {
    const fractionSection = section.slice(decimalIndex + 1);
    if (!/[0#?]/.test(fractionSection)) {
      decimalIndex = -1;
    }
  }
  const decimalSection = decimalIndex >= 0 ? section.slice(decimalIndex + 1) : '';
  const fractionDigits = decimalSection.replace(/[^0#?]/g, '').length;
  const integerSection = decimalIndex >= 0 ? section.slice(0, decimalIndex) : section;
  const useGrouping = /[,.\s\u00a0\u202f]/.test(integerSection) && integerSection.length > 0;

  return {
    fractionDigits,
    percent,
    useGrouping,
    literalPrefix: prefix || undefined,
    literalSuffix: suffix || undefined,
  };
};

const formatNumber = (value: number, info: NumberFormatInfo, locale: string): string => {
  const adjusted = info.percent ? value * 100 : value;
  const { decimal, group } = getLocaleNumberInfo(locale);
  const absValue = Math.abs(adjusted);

  const fixed = absValue.toFixed(info.fractionDigits);
  const [integerPart, fractionPart] = fixed.split('.');
  const grouped = info.useGrouping ? integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, group) : integerPart;
  const fraction = info.fractionDigits > 0 ? `${decimal}${fractionPart}` : '';
  const signed = adjusted < 0 ? '-' : '';

  let result = `${signed}${grouped}${fraction}`;

  if (info.literalPrefix) {
    result = `${info.literalPrefix}${result}`;
  }
  if (info.literalSuffix) {
    result = `${result}${info.literalSuffix}`;
  }

  return normalizeSpaces(result);
};

const padNumber = (value: number, length: number): string => {
  const str = String(value);
  return str.length >= length ? str : `${'0'.repeat(length - str.length)}${str}`;
};

const formatDatePart = (value: Date, token: string, locale: string): string => {
  switch (token) {
    case 'yyyy':
      return String(value.getFullYear());
    case 'yy':
      return padNumber(value.getFullYear() % 100, 2);
    case 'mmmm':
      return value.toLocaleString(locale, { month: 'long' });
    case 'mmm':
      return value.toLocaleString(locale, { month: 'short' });
    case 'mm':
      return padNumber(value.getMonth() + 1, 2);
    case 'm':
      return String(value.getMonth() + 1);
    case 'dd':
      return padNumber(value.getDate(), 2);
    case 'd':
      return String(value.getDate());
    case 'hh': {
      const hours = value.getHours();
      return padNumber(hours, 2);
    }
    case 'h':
      return String(value.getHours());
    case 'min2':
      return padNumber(value.getMinutes(), 2);
    case 'min1':
      return String(value.getMinutes());
    case 'ss':
      return padNumber(value.getSeconds(), 2);
    case 's':
      return String(value.getSeconds());
    default:
      return token;
  }
};

const tokenizeDateFormat = (format: string): string[] => {
  const tokens: string[] = [];
  let i = 0;
  while (i < format.length) {
    const ch = format[i];
    if (ch === '"') {
      let literal = '';
      i++;
      while (i < format.length && format[i] !== '"') {
        literal += format[i];
        i++;
      }
      i++;
      if (literal) tokens.push(literal);
      continue;
    }
    if (ch === '\\' && i + 1 < format.length) {
      tokens.push(format[i + 1]);
      i += 2;
      continue;
    }
    if (ch === '[') {
      const end = format.indexOf(']', i + 1);
      if (end !== -1) {
        i = end + 1;
        continue;
      }
    }

    const lower = format.slice(i).toLowerCase();
    const match = ['yyyy', 'yy', 'mmmm', 'mmm', 'mm', 'm', 'dd', 'd', 'hh', 'h', 'ss', 's'].find((t) =>
      lower.startsWith(t),
    );
    if (match) {
      if (match === 'm' || match === 'mm') {
        let j = i - 1;
        let previousChar = '';
        while (j >= 0 && previousChar === '') {
          const candidate = format[j];
          if (candidate && candidate !== ' ') {
            previousChar = candidate;
          }
          j--;
        }
        const isMinute = previousChar === 'h' || previousChar === 'H' || previousChar === ':';
        if (isMinute) {
          tokens.push(match === 'mm' ? 'min2' : 'min1');
          i += match.length;
          continue;
        }
      }

      tokens.push(match);
      i += match.length;
      continue;
    }

    tokens.push(ch);
    i++;
  }
  return tokens;
};

const isDateFormat = (format: string): boolean => {
  const lowered = format.toLowerCase();
  return /[ymdhss]/.test(lowered);
};

const formatDate = (value: Date, format: string, locale: string): string => {
  const tokens = tokenizeDateFormat(format);
  return tokens.map((token) => formatDatePart(value, token, locale)).join('');
};

export const formatCellValue = (
  value: number | Date,
  style: CellStyle | undefined,
  locale?: string,
): string | null => {
  const numberFormat = style?.numberFormat;
  if (!numberFormat) return null;

  const normalizedLocale = locale || DEFAULT_LOCALE;
  const sections = splitFormatSections(numberFormat);
  const hasNegativeSection = sections.length > 1;
  const section = value instanceof Date ? sections[0] : value < 0 ? sections[1] ?? sections[0] : sections[0];

  if (value instanceof Date && isDateFormat(section)) {
    return formatDate(value, section, normalizedLocale);
  }

  if (typeof value === 'number') {
    const info = parseNumberFormat(section);
    if (!info) return null;
    const numericValue = value < 0 && hasNegativeSection ? Math.abs(value) : value;
    return formatNumber(numericValue, info, normalizedLocale);
  }

  return null;
};
