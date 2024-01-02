import { CellStyle, NullableCell, SheetData } from "./sheetData";
import { expandRange } from "./utils";

export type ColWidth = {
  min: number;
  max: number;
  width: number;
};

export type ColStyle = {
  min: number;
  max: number;
  style: CellStyle;
};

export type Col = ColWidth | ColStyle;

type CombinedCol = {
  min: number;
  max: number;
  width?: number;
  style?: CellStyle;
};

// transform the same min and max into CombinedCol
export function combineColProps(cols: Col[]): CombinedCol[] {
  const combinedCols: CombinedCol[] = [];
  for (const col of cols) {
    const found = combinedCols.find(
      (c) => c.min === col.min && c.max === col.max
    );
    if (found) {
      if ("width" in col) {
        found.width = col.width;
      } else if ("style" in col) {
        found.style = col.style;
      }
      continue;
    }

    const newCombinedCol: CombinedCol = {
      min: col.min,
      max: col.max,
    };
    if ("width" in col) {
      newCombinedCol.width = col.width;
    }
    if ("style" in col) {
      newCombinedCol.style = col.style;
    }
  }

  return combinedCols;
}

export type Row = {
  index: number;
  height: number;
};

export type MergeCell = {
  /**
   * e.g. "A2:B4"
   */
  ref: string;
};

export type FreezePane = {
  type: "column" | "row";
  split: number;
};

export class Worksheet {
  private _name: string;
  private _sheetData: SheetData = [];
  private _cols: Col[] = [];
  private _rows: Row[] = [];
  private _mergeCells: MergeCell[] = [];
  private _freezePane: FreezePane | null = null;

  constructor(name: string) {
    this._name = name;
  }

  get name() {
    return this._name;
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

  setColWidth(col: ColWidth) {
    // TODO: validate col
    this._cols.push(col);
  }

  setColStyle(colStyle: ColStyle) {
    this._cols.push(colStyle);
  }

  setRowHeight(row: Row) {
    this._rows.push(row);
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
