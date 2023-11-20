import { Cell } from "./workbook";

export class SheetData {
  private data: Cell[][] = [];
  constructor() {}

  get rowsLength() {
    return this.data.length;
  }

  getCell(rowIndex: number, colIndex: number) {
    if (!this.data[rowIndex]) {
      const diff = rowIndex - this.data.length + 1;
      for (let i = 0; i < diff; i++) {
        this.data.push([]);
      }
    }

    const rows = this.data[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push({ value: "" });
      }
    }
    return rows[colIndex];
  }
}
