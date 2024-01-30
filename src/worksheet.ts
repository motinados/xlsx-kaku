import { DEFAULT_ROW_HEIGHT, RowProps } from "./row";
import { CellStyle, NullableCell, SheetData } from "./sheetData";
import { expandRange } from "./utils";

/**
 * The value is the same as the one in files created with Online Excel.
 * Changing this value will result in differences in integration tests.
 */
export const DEFAULT_COL_WIDTH = 9;

export type ColProps = {
  index: number;
  width?: number;
  style?: CellStyle;
};

export type MergeCell = {
  /**
   * e.g. "A2:B4"
   */
  ref: string;
};

export type FreezePane = {
  target: "column" | "row";
  split: number;
};

export type WorksheetProps = {
  defaultColWidth?: number;
  defaultRowHeight?: number;
};

type RequiredWorksheetProps = Required<WorksheetProps>;

export class Worksheet {
  private _name: string;
  private _props: RequiredWorksheetProps;
  private _sheetData: SheetData = [];
  private _cols = new Map<number, ColProps>();
  private _rows = new Map<number, RowProps>();
  private _mergeCells: MergeCell[] = [];
  private _freezePane: FreezePane | null = null;

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
    return this._mergeCells;
  }

  get freezePane() {
    return this._freezePane;
  }

  private getCell(rowIndex: number, colIndex: number): NullableCell {
    const rows = this._sheetData[rowIndex];
    if (!rows) {
      return null;
    }

    return rows[colIndex] || null;
  }

  // TODO: Cells that have been merged cannot be set.
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

  setColProps(colProps: ColProps) {
    this._cols.set(colProps.index, colProps);
  }

  setRowProps(row: RowProps) {
    this._rows.set(row.index, row);
  }

  setMergeCell(mergeCell: MergeCell) {
    // Within the range to be merged, cells are set with the type of "merged".
    const addresses = expandRange(mergeCell.ref);
    for (let i = 0; i < addresses.length; i++) {
      const address = addresses[i];
      if (address) {
        const [colIndex, rowIndex] = address;

        // If the cell is not set, set it as empty string.
        if (i === 0) {
          const cell = this.getCell(rowIndex, colIndex);
          if (!cell) {
            this.setCell(rowIndex, colIndex, { type: "string", value: "" });
          }
        } else {
          this.setCell(rowIndex, colIndex, { type: "merged" });
        }
      }
    }

    this._mergeCells.push(mergeCell);
  }

  setFreezePane(freezePane: FreezePane) {
    this._freezePane = freezePane;
  }
}
