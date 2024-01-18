function sortObjectKeys(obj: Record<string, any>): Record<string, any> {
  return Object.keys(obj)
    .sort()
    .reduce((sortedObj, key) => {
      sortedObj[key] = obj[key];
      return sortedObj;
    }, {} as Record<string, any>);
}

export function stringifySorted(obj: Record<string, any>): string {
  const sortedObj = sortObjectKeys(obj);
  return JSON.stringify(sortedObj);
}

/**
 * 'A1' => ['A', 1]
 */
export function devideAddress(address: string): [string, number] {
  const column = address.match(/[A-Z]+/g)![0];
  const row = address.match(/[0-9]+/g)![0];
  return [column, parseInt(row, 10)];
}

/**
 * e.g. "A" => 0
 */
export function convColNameToColIndex(colName: string): number {
  let sum = 0;
  for (let i = 0; i < colName.length; i++) {
    sum *= 26;
    sum += colName.charCodeAt(i) - "A".charCodeAt(0) + 1;
  }
  return sum - 1;
}

/**
 * e.g. 0 => "A"
 */
export function convColIndexToColName(colIndex: number): string {
  let str = "";
  while (colIndex >= 0) {
    str = String.fromCharCode((colIndex % 26) + "A".charCodeAt(0)) + str;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return str;
}

/**
 * e.g. "A1:B2" => [[0, 0], [0, 1], [1, 0], [1, 1]]
 * @param range e.g. "A1:B2"
 * @returns
 */
export function expandRange(range: string): [number, number][] {
  const [start, end] = range.split(":");
  if (!start || !end) {
    throw new Error("invalid range");
  }

  // These are not index but number. index is number - 1
  const [startColumn, startRow] = devideAddress(start);
  const [endColumn, endRow] = devideAddress(end);
  const startColumnNum = convColNameToColIndex(startColumn);
  const endColumnNum = convColNameToColIndex(endColumn);
  const startRowNum = startRow;
  const endRowNum = endRow;

  const result: [number, number][] = [];
  for (let i = startColumnNum; i <= endColumnNum; i++) {
    for (let j = startRowNum; j <= endRowNum; j++) {
      result.push([i, j - 1]); // return index
    }
  }

  return result;
}

/**
 * column name is in range
 * @param colName target column name. e.g. "A"
 * @param min min column number. It is 1 in the case of "A:C".
 * @param max max column number. It is 3 in the case of "A:C".
 * @returns
 */
export function isInRange(colName: string, min: number, max: number) {
  const columnNumber = convColNameToColIndex(colName) + 1;
  return columnNumber >= min && columnNumber <= max;
}

/**
 * if address contains "!", it means that the address has sheet name.
 * @param address e.g. "Sheet1!A1"
 * @returns
 */
export function hasSheetName(address: string): boolean {
  return address.indexOf("!") !== -1;
}
