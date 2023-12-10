import { Border } from "./borders";
import { Fill } from "./fills";
import { Font } from "./fonts";
import { NumberFormat } from "./numberFormats";

type CellStyle = {
  font?: Font;
  fill?: Fill;
  border?: Border;
  numberFormat?: NumberFormat;
};

export type Cell =
  | {
      type: "string";
      value: string;
      style?: CellStyle;
    }
  | {
      type: "number";
      value: number;
      style?: CellStyle;
    }
  | {
      type: "date";
      value: string;
      style?: CellStyle;
    }
  | {
      type: "hyperlink";
      value: string;
      style?: CellStyle;
    };

export type NullableCell = Cell | null;

export type Row = NullableCell[];

export type SheetData = Row[];
