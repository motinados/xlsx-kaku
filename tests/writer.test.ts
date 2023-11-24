import { NullableCell } from "../src/cell";
import {
  cellToString,
  findFirstNonNullCell,
  findLastNonNullCell,
  getSpans,
  rowToString,
  tableToString,
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

  test("getSpans", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "name" },
      { type: "string", value: "age" },
    ];
    const spans = getSpans(row)!;
    expect(spans.startNumber).toBe(3);
    expect(spans.endNumber).toBe(4);
  });

  test("cellToString", () => {
    const cell: NonNullable<NullableCell> = {
      type: "number",
      value: 15,
    };
    const result = cellToString(cell, 2, 0);
    expect(result).toBe(`<c r="C1"><v>15</v></c>`);
  });

  test("rowToString", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "number", value: 15 },
      { type: "number", value: 23 },
    ];
    const result = rowToString(row, 0);
    expect(result).toBe(
      `<row r="1" spans="3:4"><c r="C1"><v>15</v></c><c r="D1"><v>23</v></c></row>`
    );
  });

  test("tableToString", () => {
    const table: NullableCell[][] = [
      [],
      [null, null, { type: "number", value: 1 }, { type: "number", value: 2 }],
      [{ type: "number", value: 3 }, { type: "number", value: 4 }, null, null],
    ];
    const result = tableToString(table);
    expect(result).toBe(
      `<sheetData><row r="2" spans="3:4"><c r="C2"><v>1</v></c><c r="D2"><v>2</v></c></row><row r="3" spans="1:2"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row></sheetData>`
    );
  });
});
