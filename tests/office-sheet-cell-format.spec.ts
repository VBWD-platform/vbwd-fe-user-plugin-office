import { describe, it, expect } from 'vitest';
import { defaultDecimalsForFormat, formatNumericCellValue } from '../src/utils/sheetCellFormat';

describe('formatNumericCellValue', () => {
  it('"general" renders an integer without decoration', () => {
    expect(formatNumericCellValue(42)).toBe('42');
  });

  it('"currency" renders with the default 2 decimals', () => {
    expect(formatNumericCellValue(1234.5, { format: 'currency' })).toBe('$1,234.50');
  });

  it('"percent" multiplies by 100 and appends "%"', () => {
    expect(formatNumericCellValue(0.256, { format: 'percent', decimals: 1 })).toBe('25.6%');
  });

  it('"number" respects an explicit decimals override', () => {
    expect(formatNumericCellValue(3.14159, { format: 'number', decimals: 2 })).toBe('3.14');
  });

  it('"date" renders a numeric Excel-style serial as a calendar date', () => {
    // Serial 1 is 1899-12-31 in the Excel/Sheets epoch (day 0 = 1899-12-30).
    const rendered = formatNumericCellValue(1, { format: 'date' });
    expect(rendered).toContain('1899');
  });

  it('the stored numeric value is never mutated by formatting — display only', () => {
    const value = 1234.5;
    formatNumericCellValue(value, { format: 'currency' });
    expect(value).toBe(1234.5);
  });
});

describe('defaultDecimalsForFormat', () => {
  it('number/currency default to 2 decimals, percent to 0, general/date to 0', () => {
    expect(defaultDecimalsForFormat('number')).toBe(2);
    expect(defaultDecimalsForFormat('currency')).toBe(2);
    expect(defaultDecimalsForFormat('percent')).toBe(0);
    expect(defaultDecimalsForFormat('general')).toBe(0);
    expect(defaultDecimalsForFormat('date')).toBe(0);
  });
});
