import { NullableCell } from "./cell";

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
