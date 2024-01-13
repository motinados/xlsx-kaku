import { CellStyle } from "./sheetData";

export type RowHeight = {
  index: number;
  height: number;
};

export type RowStyle = {
  index: number;
  style: CellStyle;
};

export type Row = RowHeight | RowStyle;

export type CombinedRow = {
  index: number;
  height?: number;
  style?: CellStyle;
};

export const DEFAULT_ROW_HEIGHT = 13.5;

export function combineRowProps(rows: Row[]): CombinedRow[] {
  const combinedRows: CombinedRow[] = [];
  for (const row of rows) {
    const found = combinedRows.find((r) => r.index === row.index);
    if (found) {
      if ("height" in row) {
        found.height = row.height;
      } else if ("style" in row) {
        found.style = row.style;
      }
      continue;
    }

    const newCombinedRow: CombinedRow = {
      index: row.index,
    };
    if ("height" in row) {
      newCombinedRow.height = row.height;
    }
    if ("style" in row) {
      newCombinedRow.style = row.style;
    }
    combinedRows.push(newCombinedRow);
  }

  return combinedRows;
}