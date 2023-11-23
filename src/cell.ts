export class Cell {
  private _type: "string" | "number" = "string";
  private _value: string | number = "";
  constructor() {}

  get type() {
    return this._type;
  }

  set value(value: string | number) {
    if (typeof value === "number") {
      this._type = "number";
    } else {
      this._type = "string";
    }
    this._value = value;
  }

  get value() {
    return this._value;
  }
}

export function convColumnToNumber(column: string): number {
  let sum = 0;
  for (let i = 0; i < column.length; i++) {
    sum *= 26;
    sum += column.charCodeAt(i) - "A".charCodeAt(0) + 1;
  }
  return sum - 1;
}

export type NullableCell =
  | {
      type: "string";
      value: string;
    }
  | {
      type: "number";
      value: number;
    }
  | {
      type: "date";
      value: string;
    }
  | null;

export class Table {
  private _rows: NullableCell[][] = [];
  constructor() {}

  get rowsLength() {
    return this._rows.length;
  }

  get table() {
    return this._rows;
  }

  getCell(rowIndex: number, colIndex: number): NullableCell {
    if (!this._rows[rowIndex]) {
      return null;
    }

    const rows = this._rows[rowIndex]!;
    if (!rows[colIndex]) {
      return null;
    }

    return rows[colIndex]!;
  }

  setCellElm(rowIndex: number, colIndex: number, cell: NullableCell) {
    if (!this._rows[rowIndex]) {
      const diff = rowIndex - this._rows.length + 1;
      for (let i = 0; i < diff; i++) {
        this._rows.push([]);
      }
    }

    const rows = this._rows[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push(null);
      }
    }

    rows[colIndex] = cell;
  }
}
