import { NullableCell, SheetData } from "./sheetData";

export type Col = {
  min: number;
  max: number;
  width: number;
};

export type Row = {
  index: number;
  height: number;
};

export class Worksheet {
  private _name: string;
  private _sheetData: SheetData = [];
  private _cols: Col[] = [];
  private _rows: Row[] = [];

  constructor(name: string) {
    this._name = name;
  }

  get name() {
    return this._name;
  }

  set sheetData(sheetData: SheetData) {
    this._sheetData = sheetData;
  }

  get sheetData() {
    return this._sheetData;
  }

  get cols() {
    return this._cols;
  }

  get rows() {
    return this._rows;
  }

  setCell(rowIndex: number, colIndex: number, cell: NullableCell) {
    if (!this._sheetData[rowIndex]) {
      const diff = rowIndex - this._sheetData.length + 1;
      for (let i = 0; i < diff; i++) {
        this._sheetData.push([]);
      }
    }

    const rows = this._sheetData[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push(null);
      }
    }

    rows[colIndex] = cell;
  }

  setColWidth(col: Col) {
    // TODO: validate col
    this._cols.push(col);
  }

  setRowHeight(row: Row) {
    this._rows.push(row);
  }
}
