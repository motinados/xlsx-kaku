import { Border } from "./borders";
import { Fill } from "./fills";
import { Font } from "./fonts";
import { NumberFormat } from "./numberFormats";

export type CellStyle = {
  font?: Font;
  fill?: Fill;
  border?: Border;
  numberFormat?: NumberFormat;
  alignment?: {
    horizontal?: "left" | "center" | "right";
    vertical?: "top" | "center" | "bottom";
    textRotation?: number;
  };
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
      text: string;
      value: string;
      style?: CellStyle;
    }
  | {
      type: "boolean";
      value: boolean;
      style?: CellStyle;
    }
  | {
      type: "formula";
      value: string;
      style?: CellStyle;
    }
  | {
      type: "merged";
      style?: CellStyle;
    };

export type NullableCell = Cell | null;

export type RowData = NullableCell[];

export type SheetData = RowData[];
