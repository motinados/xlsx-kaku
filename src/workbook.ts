import { genXlsx, genXlsxSync } from "./writer";
import { Worksheet, WorksheetProps } from "./worksheet";

export class Workbook {
  private _worksheets: Worksheet[] = [];

  addWorksheet(sheetName: string, props?: WorksheetProps) {
    if (this._worksheets.some((ws) => ws.name === sheetName)) {
      throw new Error(`Worksheet name "${sheetName}" is already used.`);
    }

    const ws = new Worksheet(sheetName, props);
    this._worksheets.push(ws);
    return ws;
  }

  getWorksheet(sheetName: string) {
    return this._worksheets.find((ws) => ws.name === sheetName);
  }

  generateXlsxSync() {
    return genXlsxSync(this._worksheets);
  }

  generateXlsx() {
    return genXlsx(this._worksheets);
  }
}
