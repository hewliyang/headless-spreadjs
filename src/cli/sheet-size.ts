import type { SpreadWorksheet } from "../types.js";

/**
 * SpreadJS defaults to 200 rows × 20 columns and silently drops writes
 * outside that range. Grow sheets before writing so commands never fail
 * silently.
 */
export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLS = 16_384;

/**
 * Ensure `sheet` has at least `neededRows` rows and `neededCols` columns.
 * Only expands; never shrinks.
 */
export function ensureSheetSize(
  sheet: SpreadWorksheet,
  neededRows: number,
  neededCols: number,
): void {
  if (sheet.getRowCount() < neededRows) {
    sheet.setRowCount(neededRows);
  }
  if (sheet.getColumnCount() < neededCols) {
    sheet.setColumnCount(neededCols);
  }
}
