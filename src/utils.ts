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

export function convColumnToNumber(column: string): number {
  let sum = 0;
  for (let i = 0; i < column.length; i++) {
    sum *= 26;
    sum += column.charCodeAt(i) - "A".charCodeAt(0) + 1;
  }
  return sum - 1;
}

export function convNumberToColumn(num: number): string {
  let str = "";
  while (num >= 0) {
    str = String.fromCharCode((num % 26) + "A".charCodeAt(0)) + str;
    num = Math.floor(num / 26) - 1;
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
  const startColumnNum = convColumnToNumber(startColumn);
  const endColumnNum = convColumnToNumber(endColumn);
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
