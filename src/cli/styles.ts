/**
 * Agent-friendly style schema ↔ SpreadJS Style conversion.
 */

import type { GCNamespace, SpreadStyle, SpreadWorksheet } from "../types.js";

export interface CellStyles {
  fontWeight?: "normal" | "bold";
  fontStyle?: "normal" | "italic";
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  backgroundColor?: string;
  horizontalAlignment?: "left" | "center" | "right";
  numberFormat?: string;
}

export interface SerializedStyle {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  backgroundColor?: string;
  hAlign?: string;
  numberFormat?: string;
}

const H_ALIGN_NAMES: Record<number, string> = {
  0: "left",
  1: "center",
  2: "right",
  3: "general",
};

// Default values to suppress from serialization
const DEFAULT_FONT_SIZE = 14.6667; // SpreadJS default ~11pt
const DEFAULT_FONT_FAMILIES = new Set(["calibri", "arial"]);
const DEFAULT_FORMATTERS = new Set(["general", "General"]);
const DEFAULT_FORE_COLORS = new Set(["Text 1 0", "windowtext", "#000000"]);

export function serializeStyle(style: SpreadStyle): SerializedStyle | null {
  const result: SerializedStyle = {};
  let hasAny = false;

  if (style.fontWeight === "bold") {
    result.bold = true;
    hasAny = true;
  }
  if (style.fontStyle === "italic") {
    result.italic = true;
    hasAny = true;
  }
  if (style.fontSize) {
    const size = parseFloat(style.fontSize);
    if (Math.abs(size - DEFAULT_FONT_SIZE) > 0.5) {
      result.fontSize = size;
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
      result.fontColor = style.foreColor;
      hasAny = true;
    }
  }
  if (style.backColor && typeof style.backColor === "string") {
    result.backgroundColor = style.backColor;
    hasAny = true;
  }
  if (style.hAlign !== undefined && style.hAlign !== null) {
    const name = H_ALIGN_NAMES[style.hAlign as number];
    if (name && name !== "general") {
      result.hAlign = name;
      hasAny = true;
    }
  }
  if (style.formatter && typeof style.formatter === "string") {
    if (!DEFAULT_FORMATTERS.has(style.formatter)) {
      result.numberFormat = style.formatter;
      hasAny = true;
    }
  }

  return hasAny ? result : null;
}

export function applyStyles(
  sheet: SpreadWorksheet,
  row: number,
  col: number,
  styles: CellStyles,
  GC: GCNamespace,
): void {
  const existing = sheet.getStyle(row, col);
  const style = existing
    ? (Object.create(Object.getPrototypeOf(existing)) as SpreadStyle)
    : new GC.Spread.Sheets.Style();

  // Copy existing properties if we cloned
  if (existing) {
    Object.assign(style, existing);
  }

  if (styles.fontWeight) {
    style.fontWeight = styles.fontWeight;
  }
  if (styles.fontStyle) {
    style.fontStyle = styles.fontStyle;
  }
  if (styles.fontSize !== undefined) {
    style.fontSize = `${styles.fontSize}px`;
  }
  if (styles.fontFamily) {
    style.fontFamily = styles.fontFamily;
  }
  if (styles.fontColor) {
    style.foreColor = styles.fontColor;
  }
  if (styles.backgroundColor) {
    style.backColor = styles.backgroundColor;
  }
  if (styles.numberFormat) {
    style.formatter = styles.numberFormat;
  }
  if (styles.horizontalAlignment) {
    const alignMap: Record<string, number> = {
      left: 0,
      center: 1,
      right: 2,
    };
    style.hAlign = alignMap[styles.horizontalAlignment] as any;
  }

  sheet.setStyle(row, col, style);
}
