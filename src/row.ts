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
