import { CellStyle } from "./sheetData";

/**
 * The value is the same as the one in files created with Online Excel.
 * Changing this value will result in differences in integration tests.
 */
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

export type ColProps = {
  index: number;
  width?: number;
  style?: CellStyle;
};
