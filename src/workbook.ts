import { Worksheet } from "./worksheet";

export type Cell = {
  value: number | string;
};

export type Row = {
  cells: Cell[];
};

export type WorksheetData = {
  rows: Row[];
};

export class Workbook {
  private sheets: Worksheet[] = [];
  constructor() {
    const sheet = new Worksheet({ sheetName: "Sheet1" });
    this.sheets.push(sheet);
  }
}
