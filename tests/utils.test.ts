import {
  convColumnToNumber,
  convNumberToColumn,
  devideAddress,
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
});
