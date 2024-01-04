import {
  convColumnToNumber,
  convNumberToColumn,
  devideAddress,
  expandRange,
  isInRange,
} from "../src/utils";

describe("utils", () => {
  test("convColumnToNumber", () => {
    expect(convColumnToNumber("A")).toBe(0);
    expect(convColumnToNumber("B")).toBe(1);
    expect(convColumnToNumber("Z")).toBe(25);
    expect(convColumnToNumber("AA")).toBe(26);
    expect(convColumnToNumber("BC")).toBe(54);
  });

  test("convNumberToColumn", () => {
    expect(convNumberToColumn(0)).toBe("A");
    expect(convNumberToColumn(1)).toBe("B");
    expect(convNumberToColumn(25)).toBe("Z");
    expect(convNumberToColumn(26)).toBe("AA");
    expect(convNumberToColumn(54)).toBe("BC");
  });

  test("devideAddress", () => {
    expect(devideAddress("A1")).toStrictEqual(["A", 1]);
    expect(devideAddress("B2")).toStrictEqual(["B", 2]);
    expect(devideAddress("Z3")).toStrictEqual(["Z", 3]);
    expect(devideAddress("AA10")).toStrictEqual(["AA", 10]);
    expect(devideAddress("BCD99")).toStrictEqual(["BCD", 99]);
  });

  test("expandRange", () => {
    expect(expandRange("A1:A1")).toStrictEqual([[0, 0]]);
    expect(expandRange("A1:A2")).toStrictEqual([
      [0, 0],
      [0, 1],
    ]);
    expect(expandRange("A1:B2")).toStrictEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
    expect(expandRange("A1:C3")).toStrictEqual([
      [0, 0],
      [0, 1],
      [0, 2],
      [1, 0],
      [1, 1],
      [1, 2],
      [2, 0],
      [2, 1],
      [2, 2],
    ]);
  });

  test("isInRange", () => {
    expect(isInRange("A", 1, 1)).toBe(true);
    expect(isInRange("A", 1, 2)).toBe(true);
    expect(isInRange("A", 2, 2)).toBe(false);

    expect(isInRange("C", 1, 1)).toBe(false);
    expect(isInRange("C", 1, 2)).toBe(false);
    expect(isInRange("C", 3, 3)).toBe(true);
    expect(isInRange("C", 3, 4)).toBe(true);
  });
});
