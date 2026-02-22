/**
 * SpreadJS-style cell style serialization and application for CLI I/O.
 */

import type { Spread } from "@mescius/spread-sheets";
import type { GCNamespace, SpreadStyle, SpreadWorksheet } from "../types.js";

export interface FontStyleObject {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface CellStyle {
  fontStyle?: FontStyleObject;
  fontSize?: string;
  fontFamily?: string;
  foreColor?: string;
  backColor?: string;
  hAlign?: Spread.Sheets.HorizontalAlign;
  formatter?: string;
}

const H_ALIGN_GENERAL = 3;

// Default values to suppress from serialization
const DEFAULT_FONT_SIZE = 14.6667; // SpreadJS default ~11pt
const DEFAULT_FONT_FAMILIES = new Set(["calibri", "arial"]);
const DEFAULT_FORMATTERS = new Set(["general", "General"]);
const DEFAULT_FORE_COLORS = new Set(["Text 1 0", "windowtext", "#000000"]);

function normalizeFontStyle(style: SpreadStyle): FontStyleObject | undefined {
  const result: FontStyleObject = {};

  const raw = (style as SpreadStyle & { fontStyle?: unknown }).fontStyle;
  if (raw && typeof raw === "object") {
    const maybe = raw as FontStyleObject;
    if (maybe.bold === true) result.bold = true;
    if (maybe.italic === true) result.italic = true;
    if (maybe.underline === true) result.underline = true;
  }

  if (style.fontWeight === "bold") {
    result.bold = true;
  }

  if (style.fontStyle === "italic") {
    result.italic = true;
  }

  const textDecoration = (style as SpreadStyle & { textDecoration?: unknown })
    .textDecoration;
  if (textDecoration !== undefined && textDecoration !== null) {
    if (typeof textDecoration === "number") {
      if (textDecoration !== 0) {
        result.underline = true;
      }
    } else if (typeof textDecoration === "string") {
      if (textDecoration !== "none" && textDecoration !== "None") {
        result.underline = true;
      }
    }
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

export function serializeStyle(style: SpreadStyle): CellStyle | null {
  const result: CellStyle = {};
  let hasAny = false;

  const fontStyle = normalizeFontStyle(style);
  if (fontStyle) {
    result.fontStyle = fontStyle;
    hasAny = true;
  }

  if (style.fontSize) {
    const size = parseFloat(style.fontSize);
    if (Number.isFinite(size) && Math.abs(size - DEFAULT_FONT_SIZE) > 0.5) {
      result.fontSize = style.fontSize;
      hasAny = true;
    }
  }

  if (style.fontFamily) {
    if (!DEFAULT_FONT_FAMILIES.has(style.fontFamily.toLowerCase())) {
      result.fontFamily = style.fontFamily;
      hasAny = true;
    }
  }

  if (style.foreColor) {
    if (!DEFAULT_FORE_COLORS.has(style.foreColor)) {
      result.foreColor = style.foreColor;
      hasAny = true;
    }
  }

  if (style.backColor && typeof style.backColor === "string") {
    result.backColor = style.backColor;
    hasAny = true;
  }

  if (style.hAlign !== undefined && style.hAlign !== null) {
    if ((style.hAlign as number) !== H_ALIGN_GENERAL) {
      result.hAlign = style.hAlign as Spread.Sheets.HorizontalAlign;
      hasAny = true;
    }
  }

  if (style.formatter && typeof style.formatter === "string") {
    if (!DEFAULT_FORMATTERS.has(style.formatter)) {
      result.formatter = style.formatter;
      hasAny = true;
    }
  }

  return hasAny ? result : null;
}

export function applyStyles(
  sheet: SpreadWorksheet,
  row: number,
  col: number,
  styles: CellStyle,
  GC: GCNamespace,
): void {
  const existing = sheet.getStyle(row, col);
  const style = existing
    ? (Object.create(Object.getPrototypeOf(existing)) as SpreadStyle)
    : new GC.Spread.Sheets.Style();

  if (existing) {
    Object.assign(style, existing);
  }

  if (styles.fontStyle) {
    if (styles.fontStyle.bold !== undefined) {
      style.fontWeight = styles.fontStyle.bold ? "bold" : "normal";
    }

    if (styles.fontStyle.italic !== undefined) {
      style.fontStyle = styles.fontStyle.italic ? "italic" : "normal";
    }

    if (styles.fontStyle.underline !== undefined) {
      const TextDecoration = GC.Spread.Sheets.TextDecorationType;
      style.textDecoration = styles.fontStyle.underline
        ? TextDecoration.underline
        : TextDecoration.none;
    }
  }

  if (styles.fontSize !== undefined) {
    style.fontSize = styles.fontSize;
  }

  if (styles.fontFamily !== undefined) {
    style.fontFamily = styles.fontFamily;
  }

  if (styles.foreColor !== undefined) {
    style.foreColor = styles.foreColor;
  }

  if (styles.backColor !== undefined) {
    style.backColor = styles.backColor;
  }

  if (styles.formatter !== undefined) {
    style.formatter = styles.formatter;
  }

  if (styles.hAlign !== undefined) {
    style.hAlign = styles.hAlign;
  }

  sheet.setStyle(row, col, style);
}
