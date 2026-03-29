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

export interface LineBorderObject {
  color?: string;
  style?: string;
}

export interface CellStyle {
  fontStyle?: FontStyleObject;
  fontSize?: string;
  fontFamily?: string;
  foreColor?: string;
  backColor?: string;
  hAlign?: Spread.Sheets.HorizontalAlign;
  vAlign?: number;
  formatter?: string;
  wordWrap?: boolean;
  textIndent?: number;
  borderLeft?: LineBorderObject;
  borderTop?: LineBorderObject;
  borderRight?: LineBorderObject;
  borderBottom?: LineBorderObject;
}

const H_ALIGN_GENERAL = 3;

const LINE_STYLE_NAMES: Record<number, string> = {
  0: "empty",
  1: "thin",
  2: "medium",
  3: "dashed",
  4: "dotted",
  5: "thick",
  6: "double",
  7: "hair",
  8: "mediumDashed",
  9: "dashDot",
  10: "mediumDashDot",
  11: "dashDotDot",
  12: "mediumDashDotDot",
  13: "slantedDashDot",
};

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

  if (style.vAlign !== undefined && style.vAlign !== null && (style.vAlign as number) !== 0) {
    result.vAlign = style.vAlign as number;
    hasAny = true;
  }

  if (style.wordWrap === true) {
    result.wordWrap = true;
    hasAny = true;
  }

  if (style.textIndent !== undefined && style.textIndent > 0) {
    result.textIndent = style.textIndent as number;
    hasAny = true;
  }

  const borderProps = ["borderLeft", "borderTop", "borderRight", "borderBottom"] as const;
  for (const prop of borderProps) {
    const border = (style as unknown as Record<string, unknown>)[prop] as { color?: string; style?: unknown } | undefined;
    if (border && typeof border === "object") {
      const borderStyle = typeof border.style === "number"
        ? LINE_STYLE_NAMES[border.style]
        : typeof border.style === "string" ? border.style : undefined;
      (result as Record<string, unknown>)[prop] = {
        color: border.color,
        ...(borderStyle && borderStyle !== "empty" ? { style: borderStyle } : {}),
      };
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

  if (styles.vAlign !== undefined) {
    style.vAlign = styles.vAlign as Spread.Sheets.VerticalAlign;
  }

  if (styles.wordWrap !== undefined) {
    style.wordWrap = styles.wordWrap;
  }

  if (styles.textIndent !== undefined) {
    style.textIndent = styles.textIndent;
  }

  const borderProps = ["borderLeft", "borderTop", "borderRight", "borderBottom"] as const;
  for (const prop of borderProps) {
    const border = styles[prop];
    if (border) {
      (style as unknown as Record<string, unknown>)[prop] = new GC.Spread.Sheets.LineBorder(
        border.color || "#000000",
        GC.Spread.Sheets.LineStyle[border.style as keyof typeof GC.Spread.Sheets.LineStyle] ?? GC.Spread.Sheets.LineStyle.thin,
      );
    }
  }

  sheet.setStyle(row, col, style);
}
