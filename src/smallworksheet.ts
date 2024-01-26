import { Worksheet } from ".";
import { Col, ColStyle, ColWidth, DEFAULT_COL_WIDTH } from "./col";
import { DEFAULT_ROW_HEIGHT, Row, RowHeight, RowStyle } from "./row";
import { NullableCell, SheetData } from "./sheetData";
import {
  FreezePane,
  RequiredWorksheetProps,
  WorksheetProps,
} from "./worksheet";

export class SmallWorksheet implements Worksheet {
  private _name: string;
  private _props: RequiredWorksheetProps;
  private _sheetData: SheetData = [];
  private _cols: Col[] = [];
  private _rows: Row[] = [];
  // private _mergeCells: MergeCell[] = [];
  private _freezePane: FreezePane | null = null;
  private _mergeCellModule = null;

  constructor(name: string, props: WorksheetProps | undefined = {}) {
    this._name = name;

    this._props = {
      defaultColWidth: props.defaultColWidth ?? DEFAULT_COL_WIDTH,
      defaultRowHeight: props.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
    };
  }

  get name() {
    return this._name;
  }

  get props() {
    return this._props;
  }

  set sheetData(sheetData: SheetData) {
    this._sheetData = sheetData;
  }

  get sheetData() {
    return this._sheetData;
  }

  get cols() {
    return this._cols;
  }

  get rows() {
    return this._rows;
  }

  get mergeCells() {
    return [];
  }

  get freezePane() {
    return this._freezePane;
  }

  get mergeCellModule() {
    return this._mergeCellModule;
  }

  getCell(rowIndex: number, colIndex: number): NullableCell {
    const rows = this._sheetData[rowIndex];
    if (!rows) {
      return null;
    }

    return rows[colIndex] || null;
  }
  setCell(rowIndex: number, colIndex: number, cell: NullableCell) {
    if (!this._sheetData[rowIndex]) {
      const diff = rowIndex - this._sheetData.length + 1;
      for (let i = 0; i < diff; i++) {
        this._sheetData.push([]);
      }
    }

    const rows = this._sheetData[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push(null);
      }
    }

    rows[colIndex] = cell;
  }

  setColWidth(col: ColWidth) {
    // TODO: validate col
    this._cols.push(col);
  }

  setColStyle(colStyle: ColStyle) {
    this._cols.push(colStyle);
  }

  setRowHeight(row: RowHeight) {
    this._rows.push(row);
  }

  setRowStyle(row: RowStyle) {
    this._rows.push(row);
  }

  setFreezePane(freezePane: FreezePane) {
    this._freezePane = freezePane;
  }
}
