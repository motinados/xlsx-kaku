import { NullableCell, convNumberToColumn } from "./cell";
import { SharedStrings } from "./sharedStrings";

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
  const sharedStrings = new SharedStrings();

  const sheetDataString = makeSheetDataXml(table, sharedStrings);
  const sharedStringsXml = makeSharedStringsXml(sharedStrings);
  return { sheetDataString, sharedStringsXml };
}

export function makeSharedStringsXml(sharedStrings: SharedStrings) {
  if (sharedStrings.count === 0) {
    return null;
  }

  let result = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.count}" uniqueCount="${sharedStrings.uniqueCount}">`;
  for (const str of sharedStrings.getValuesInOrder()) {
    result += `<si><t>${str}</t></si>`;
  }
  result += `</sst>`;
  return result;
}

export function makeSheetDataXml(
  table: NullableCell[][],
  sharedStrings: SharedStrings
) {
  const { startNumber, endNumber } = getSpansFromTable(table);

  let result = `<sheetData>`;
  let rowIndex = 0;
  for (const row of table) {
    const str = rowToString(
      row,
      rowIndex,
      startNumber,
      endNumber,
      sharedStrings
    );
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
  rowIndex: number,
  startNumber: number,
  endNumber: number,
  sharedStrings: SharedStrings
): string | null {
  if (row.length === 0) {
    return null;
  }

  const rowNumber = rowIndex + 1;
  let result = `<row r="${rowNumber}" spans="${startNumber}:${endNumber}">`;

  let columnIndex = 0;
  for (const cell of row) {
    if (cell !== null) {
      result += cellToString(cell, columnIndex, rowIndex, sharedStrings);
    }

    columnIndex++;
  }

  result += `</row>`;
  return result;
}

export function getSpansFromTable(table: NullableCell[][]) {
  const all = table
    .map((row) => {
      const spans = getSpans(row);
      if (spans === null) {
        return null;
      }
      return spans;
    })
    .filter((row) => row !== null) as {
    startNumber: number;
    endNumber: number;
  }[];
  const minStartNumber = Math.min(...all.map((row) => row.startNumber));
  const maxEndNumber = Math.max(...all.map((row) => row.endNumber));
  return { startNumber: minStartNumber, endNumber: maxEndNumber };
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
  rowIndex: number,
  sharedStrings: SharedStrings
) {
  const rowNumber = rowIndex + 1;
  const column = convNumberToColumn(columnIndex);
  switch (cell.type) {
    case "number": {
      return `<c r="${column}${rowNumber}"><v>${cell.value}</v></c>`;
    }
    case "string": {
      const index = sharedStrings.getIndex(cell.value);
      return `<c r="${column}${rowNumber}" t="s"><v>${index}</v></c>`;
    }
    default: {
      throw new Error(`not implemented: ${cell.type}`);
    }
  }
}
