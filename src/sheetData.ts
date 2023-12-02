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

export class SheetData {
  private _rows: NullableCell[][] = [];
  constructor() {}

  get rowsLength() {
    return this._rows.length;
  }

  get table() {
    return this._rows;
  }

  getCell(rowIndex: number, colIndex: number): NullableCell {
    if (!this._rows[rowIndex]) {
      return null;
    }

    const rows = this._rows[rowIndex]!;
    if (!rows[colIndex]) {
      return null;
    }

    return rows[colIndex]!;
  }

  setCell(rowIndex: number, colIndex: number, cell: NullableCell) {
    if (!this._rows[rowIndex]) {
      const diff = rowIndex - this._rows.length + 1;
      for (let i = 0; i < diff; i++) {
        this._rows.push([]);
      }
    }

    const rows = this._rows[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push(null);
      }
    }

    rows[colIndex] = cell;
  }
}

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
