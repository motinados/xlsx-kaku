import { Worksheet } from "./worksheet";
import { writeXlsx } from "./writer";

export class Workbook {
  private _worksheets: Worksheet[] = [];

  addWorksheet(sheetName: string) {
    if (this._worksheets.some((ws) => ws.name === sheetName)) {
      throw new Error(`Worksheet name "${sheetName}" is already used.`);
    }

    const ws = new Worksheet(sheetName);
    this._worksheets.push(ws);
    return ws;
  }

  getWorksheet(sheetName: string) {
    return this._worksheets.find((ws) => ws.name === sheetName);
  }

  async save(filepath: string) {
    await writeXlsx(filepath, this._worksheets);
  }
}
