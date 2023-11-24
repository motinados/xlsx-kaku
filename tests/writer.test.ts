import { NullableCell } from "../src/cell";
import {
  cellToString,
  findFirstNonNullCell,
  findLastNonNullCell,
  rowToString,
} from "../src/writer";

describe("Writer", () => {
  test("findFirstNonNullCell", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "name" },
      { type: "string", value: "age" },
    ];
    const { firstNonNullCell, index } = findFirstNonNullCell(row);
    expect(firstNonNullCell).toEqual({ type: "string", value: "name" });
    expect(index).toBe(2);
  });

  test("findLastNonNullCell", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "name" },
      { type: "string", value: "age" },
    ];
    const { lastNonNullCell, index } = findLastNonNullCell(row);
    expect(lastNonNullCell).toEqual({ type: "string", value: "age" });
    expect(index).toBe(3);
  });

  test("cellToString", () => {
    const cell: NonNullable<NullableCell> = {
      type: "number",
      value: 15,
    };
    const result = cellToString(cell, 2, 1);
    expect(result).toBe(`<c r="C1"><v>15</v></c>`);
  });

  test("rowToString", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "number", value: 15 },
      { type: "number", value: 23 },
    ];
    const result = rowToString(row, 1);
    expect(result).toBe(
      `<row r="1" spans="2:3"><c r="C1"><v>15</v></c><c r="D1"><v>23</v></c></row>`
    );
  });
});
