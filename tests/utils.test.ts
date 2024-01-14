import {
  convColNameToColIndex,
  convColIndexToColName,
  devideAddress,
  expandRange,
  isInRange,
} from "../src/utils";

describe("utils", () => {
  test("convColumnToNumber", () => {
    expect(convColNameToColIndex("A")).toBe(0);
    expect(convColNameToColIndex("B")).toBe(1);
    expect(convColNameToColIndex("Z")).toBe(25);
    expect(convColNameToColIndex("AA")).toBe(26);
    expect(convColNameToColIndex("BC")).toBe(54);
  });

  test("convNumberToColumn", () => {
    expect(convColIndexToColName(0)).toBe("A");
    expect(convColIndexToColName(1)).toBe("B");
    expect(convColIndexToColName(25)).toBe("Z");
    expect(convColIndexToColName(26)).toBe("AA");
    expect(convColIndexToColName(54)).toBe("BC");
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
