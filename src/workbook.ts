import { genXlsx, genXlsxSync } from "./writer";
import {
  Worksheet,
  WorksheetProps,
  WorksheetS,
  WorksheetType,
} from "./worksheet";
import { ImageStore } from "./imageStore";

/**
 * Standard Workbook class
 */
export class Workbook {
  private _worksheets: WorksheetType[] = [];
  private _imageStore: ImageStore = new ImageStore();

  addWorksheet(sheetName: string, props?: WorksheetProps) {
    if (this._worksheets.some((ws) => ws.name === sheetName)) {
      throw new Error(`Worksheet name "${sheetName}" is already used.`);
    }

    const ws = new Worksheet(sheetName, this._imageStore, props);
    this._worksheets.push(ws);
    return ws;
  }

  getWorksheet(sheetName: string) {
    return this._worksheets.find((ws) => ws.name === sheetName);
  }

  generateXlsxSync() {
    return genXlsxSync(this._worksheets, this._imageStore);
  }

  generateXlsx() {
    return genXlsx(this._worksheets, this._imageStore);
  }
}

/**
 * Simplified Workbook class
 */
export class WorkbookS {
  private _worksheets: WorksheetType[] = [];

  addWorksheet(sheetName: string, props?: WorksheetProps) {
    if (this._worksheets.some((ws) => ws.name === sheetName)) {
      throw new Error(`Worksheet name "${sheetName}" is already used.`);
    }

    const ws = new WorksheetS(sheetName, props);
    this._worksheets.push(ws);
    return ws;
  }

  getWorksheet(sheetName: string) {
    return this._worksheets.find((ws) => ws.name === sheetName);
  }

  generateXlsxSync() {
    return genXlsxSync(this._worksheets, null);
  }

  generateXlsx() {
    return genXlsx(this._worksheets, null);
  }
}
