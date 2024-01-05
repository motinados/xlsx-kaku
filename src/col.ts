import { CellStyle } from "./sheetData";

export const DEFAULT_COL_WIDTH = 9;

export type ColWidth = {
  startIndex: number;
  endIndex: number;
  width: number;
};

export type ColStyle = {
  startIndex: number;
  endIndex: number;
  style: CellStyle;
};

export type Col = ColWidth | ColStyle;

export type CombinedCol = {
  startIndex: number;
  endIndex: number;
  width?: number;
  style?: CellStyle;
};

// transform the same min and max into CombinedCol
export function combineColProps(cols: Col[]): CombinedCol[] {
  const combinedCols: CombinedCol[] = [];
  for (const col of cols) {
    const found = combinedCols.find(
      (c) => c.startIndex === col.startIndex && c.endIndex === col.endIndex
    );
    if (found) {
      if ("width" in col) {
        found.width = col.width;
      } else if ("style" in col) {
        found.style = col.style;
      }
      continue;
    }

    const newCombinedCol: CombinedCol = {
      startIndex: col.startIndex,
      endIndex: col.endIndex,
    };
    if ("width" in col) {
      newCombinedCol.width = col.width;
    }
    if ("style" in col) {
      newCombinedCol.style = col.style;
    }
    combinedCols.push(newCombinedCol);
  }

  return combinedCols;
}
