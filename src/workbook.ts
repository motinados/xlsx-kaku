import { Worksheet } from "./worksheet";
import { writeXlsx } from "./writer";

export class Workbook {
  private _worksheets: Worksheet[] = [];

  addWorksheet(name: string) {
    const ws = new Worksheet(name);
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
