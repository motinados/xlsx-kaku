import { CellStyle } from "./sheetData";

export type RowProps = {
  index: number;
  height?: number;
  style?: CellStyle;
};

export const DEFAULT_ROW_HEIGHT = 13.5;
