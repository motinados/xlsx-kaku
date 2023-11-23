import { NullableCell } from "../src/cell";
import { findFirstNonNullCell } from "../src/writer";

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
});
