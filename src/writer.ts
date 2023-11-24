import { NullableCell, convNumberToColumn } from "./cell";

export function findFirstNonNullCell(row: NullableCell[]) {
  let index = -1;
  let firstNonNullCell = null;
  for (let i = 0; i < row.length; i++) {
    if (row[i] !== null) {
      firstNonNullCell = row[i]!;
      index = i;
      break;
    }
  }
  return { firstNonNullCell, index };
}

/**
 *  [null, null, null, nonnull, null] => index is 3
 */
export function findLastNonNullCell(row: NullableCell[]) {
  let index = -1;
  let lastNonNullCell = null;
  for (let i = row.length - 1; i >= 0; i--) {
    if (row[i] !== null) {
      lastNonNullCell = row[i]!;
      index = i;
      break;
    }
  }
  return { lastNonNullCell, index };
}

/**
 * <row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
 */
export function rowToString(row: NullableCell[], rowIndex: number) {
  const first = findFirstNonNullCell(row);
  if (first === undefined || first === null) {
    throw new Error("row is empty");
  }

  const last = findLastNonNullCell(row);
  if (last === undefined) {
    throw new Error("row is empty");
  }

  const startSpan = first.index;
  const lastSpan = last.index;

  let result = `<row r="${rowIndex}" spans="${startSpan}:${lastSpan}">`;

  let columnIndex = 0;
  for (const cell of row) {
    if (cell !== null) {
      result += cellToString(cell, columnIndex, rowIndex);
    }

    columnIndex++;
  }

  result += `</row>`;
  return result;
}

export function cellToString(
  cell: NonNullable<NullableCell>,
  columnIndex: number,
  rowIndex: number
) {
  const column = convNumberToColumn(columnIndex);
  switch (cell.type) {
    case "number": {
      return `<c r="${column}${rowIndex}"><v>${cell.value}</v></c>`;
    }
    default: {
      throw new Error(`not implemented: ${cell.type}`);
    }
  }
}
