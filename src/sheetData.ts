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

/**
 * 'A1' => ['A', 1]
 */
export function devideAddress(address: string): [string, number] {
  const column = address.match(/[A-Z]+/g)![0];
  const row = address.match(/[0-9]+/g)![0];
  return [column, parseInt(row, 10)];
}

export function convColumnToNumber(column: string): number {
  let sum = 0;
  for (let i = 0; i < column.length; i++) {
    sum *= 26;
    sum += column.charCodeAt(i) - "A".charCodeAt(0) + 1;
  }
  return sum - 1;
}

export function convNumberToColumn(num: number): string {
  let str = "";
  while (num >= 0) {
    str = String.fromCharCode((num % 26) + "A".charCodeAt(0)) + str;
    num = Math.floor(num / 26) - 1;
  }
  return str;
}
