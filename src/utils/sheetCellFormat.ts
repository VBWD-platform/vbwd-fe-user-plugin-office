// DISPLAY-ONLY number formatting for the S147-4 drag/toolbar slice — mirrors
// the backend's `sheet_content.normalize_cell_style` "format" values exactly
// (general/number/currency/percent/date). Formatting NEVER changes the
// stored cell value: a currency-formatted `42` is still the number `42` on
// the wire and in the engine, only its rendered text differs — getting this
// wrong would silently corrupt every total the sheet computes.
import type { OfficeSheetCellStyle, OfficeSheetNumberFormat } from '../api/officeApi';

//: Excel/Sheets' own epoch for a numeric "date serial" — day 0 is
//: 1899-12-30 (not 1900-01-01; Excel's serial numbering famously treats 1900
//: as a leap year, so backdating one extra day is the standard correction).
//: This is ONLY used to render a plain number under `format: "date"`; a cell
//: whose value already carries the engine's own `{t:'date', v: iso}` wrapper
//: never reaches this function (see `OfficeSheetEditor.vue`'s `displayValue`).
const EXCEL_SERIAL_DAY_ZERO_UTC_MS = Date.UTC(1899, 11, 30);
const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;

const DEFAULT_DECIMALS_BY_FORMAT: Partial<Record<OfficeSheetNumberFormat, number>> = {
  number: 2,
  currency: 2,
  percent: 0,
};

export const MAXIMUM_STYLE_DECIMAL_PLACES = 10;

export function defaultDecimalsForFormat(format: OfficeSheetNumberFormat): number {
  return DEFAULT_DECIMALS_BY_FORMAT[format] ?? 0;
}

function dateSerialToDisplayText(serial: number): string {
  const milliseconds = EXCEL_SERIAL_DAY_ZERO_UTC_MS + serial * MILLISECONDS_PER_DAY;
  return new Date(milliseconds).toLocaleDateString();
}

/** Render a NUMERIC cell value for display per its style's `format` — the
 * one place every number-formatting toolbar action (currency/percent/
 * number/date) is actually applied. Text/boolean/error/date-wrapper cell
 * values never pass through here (they have their own display rules in
 * `OfficeSheetEditor.vue`). */
export function formatNumericCellValue(value: number, style?: OfficeSheetCellStyle): string {
  const format = style?.format ?? 'general';
  if (format === 'general') {
    return Number.isInteger(value) ? String(value) : value.toLocaleString();
  }
  if (format === 'date') return dateSerialToDisplayText(value);

  const decimals = style?.decimals ?? defaultDecimalsForFormat(format);
  if (format === 'percent') return `${(value * 100).toFixed(decimals)}%`;
  if (format === 'currency') {
    return new Intl.NumberFormat(undefined, {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(value);
  }
  return value.toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}
