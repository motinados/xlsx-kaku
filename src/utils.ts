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
