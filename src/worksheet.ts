import { Row } from "./workbook";

export class Worksheet {
  private rows: Row[] = [];
  constructor() {}

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
