import { NullableCell } from "./sheetData";

export class Worksheet {
  private _name: string;
  private _rows: NullableCell[][] = [];
  constructor(name: string) {
    this._name = name;
  }

  get name() {
    return this._name;
  }

  set sheetData(rows: NullableCell[][]) {
    this._rows = rows;
  }

  get sheetData() {
    return this._rows;
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
