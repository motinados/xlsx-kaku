import { CellStyle } from "./sheetData";

export const DEFAULT_COL_WIDTH = 9;

export type ColWidth = {
  min: number;
  max: number;
  width: number;
};

export type ColStyle = {
  min: number;
  max: number;
  style: CellStyle;
};

export type Col = ColWidth | ColStyle;

type CombinedCol = {
  min: number;
  max: number;
  width?: number;
  style?: CellStyle;
};

// transform the same min and max into CombinedCol
export function combineColProps(cols: Col[]): CombinedCol[] {
  const combinedCols: CombinedCol[] = [];
  for (const col of cols) {
    const found = combinedCols.find(
      (c) => c.min === col.min && c.max === col.max
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
      min: col.min,
      max: col.max,
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
