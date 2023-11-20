import { Row } from "./workbook";

export class Rows {
  private rows: Row[] = [];
  constructor() {}

  get length() {
    return this.rows.length;
  }

  getRow(index: number) {
    if (!this.rows[index]) {
      const diff = index - this.rows.length + 1;
      for (let i = 0; i < diff; i++) {
        this.rows.push({ cells: [] });
      }
    }
    return this.rows[index];
  }
}
