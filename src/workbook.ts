import { genXlsx } from "./writer";
import { BasicWorksheet, Worksheet, WorksheetProps } from "./worksheet";

export class Workbook {
  private _worksheets: Worksheet[] = [];

  addWorksheet(sheetName: string, props?: WorksheetProps): Worksheet {
    if (this._worksheets.some((ws) => ws.name === sheetName)) {
      throw new Error(`Worksheet name "${sheetName}" is already used.`);
    }

    const ws = new BasicWorksheet(sheetName, props);
    this._worksheets.push(ws);
    return ws;
  }

  getWorksheet(sheetName: string) {
    return this._worksheets.find((ws) => ws.name === sheetName);
  }

  generateXlsx() {
    return genXlsx(this._worksheets);
  }
}
