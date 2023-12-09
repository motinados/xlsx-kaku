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
}
