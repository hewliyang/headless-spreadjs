import "@mescius/spread-sheets";

declare module "@mescius/spread-sheets" {
  namespace Spread.Sheets {
    interface Worksheet {
      setFormatter(
        row: number,
        col: number,
        value: string | GC.Spread.Formatter.FormatterBase,
        sheetArea?: GC.Spread.Sheets.SheetArea,
      ): void;
    }
  }
}
