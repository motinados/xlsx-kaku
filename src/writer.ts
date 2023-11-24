import { NullableCell, convNumberToColumn } from "./cell";

export function findFirstNonNullCell(row: NullableCell[]) {
  let index = 0;
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
  let index = 0;
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

export function tableToString(table: NullableCell[][]) {
  let result = `<sheetData>`;
  let rowIndex = 0;
  for (const row of table) {
    const str = rowToString(row, rowIndex);
    if (str !== null) {
      result += str;
    }
    rowIndex++;
  }
  result += `</sheetData>`;
  return result;
}

/**
 * <row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
 */
export function rowToString(
  row: NullableCell[],
  rowIndex: number
): string | null {
  if (row.length === 0) {
    return null;
  }

  const spans = getSpans(row);
  if (spans === null) {
    return null;
  }

  const { startNumber, endNumber } = spans;
  const rowNumber = rowIndex + 1;
  let result = `<row r="${rowNumber}" spans="${startNumber}:${endNumber}">`;

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

export function getSpans(row: NullableCell[]) {
  const first = findFirstNonNullCell(row);
  if (first === undefined || first === null) {
    return null;
  }

  const last = findLastNonNullCell(row);
  if (last === undefined) {
    return null;
  }

  const startNumber = first.index + 1;
  const endNumber = last.index + 1;

  return { startNumber, endNumber };
}

export function cellToString(
  cell: NonNullable<NullableCell>,
  columnIndex: number,
  rowIndex: number
) {
  const rowNumber = rowIndex + 1;
  const column = convNumberToColumn(columnIndex);
  switch (cell.type) {
    case "number": {
      return `<c r="${column}${rowNumber}"><v>${cell.value}</v></c>`;
    }
    default: {
      throw new Error(`not implemented: ${cell.type}`);
    }
  }
}
