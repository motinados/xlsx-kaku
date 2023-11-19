import { Row } from "./workbook";

export class Worksheet {
  private _sheetName: string = "";
  private rows: Row[] = [];

  constructor({ sheetName }: { sheetName: string }) {
    this.sheetName = sheetName;
  }

  set sheetName(name: string) {
    this._sheetName = name;
  }

  get sheetName() {
    return this._sheetName;
  }

  addRow(row: Row) {
    this.rows.push(row);
  }

  getRow(index: number) {
    if (index >= this.rows.length) {
      throw new Error("Index out of bounds");
    }
    return this.rows[index];
  }

  getRows() {
    return this.rows;
  }
}
