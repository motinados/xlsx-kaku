import {
  SheetData,
  convColumnToNumber,
  convNumberToColumn,
  devideAddress,
} from "../src/sheetData";

describe("SheetData", () => {
  test("should be able to create a cell", () => {
    const table = new SheetData();
    expect(table).toBeInstanceOf(SheetData);
  });

  test("should be able to get a cell", () => {
    const table = new SheetData();
    const cell = table.getCell(0, 0);
    expect(cell).toBeNull();
  });

  test("should be able to set a cell", () => {
    const table = new SheetData();
    const cell = table.getCell(0, 0);
    expect(cell).toBeNull();

    table.setCell(0, 0, { type: "string", value: "hello" });
    const cell2 = table.getCell(0, 0);
    expect(cell2).toEqual({ type: "string", value: "hello" });

    table.setCell(0, 0, { type: "number", value: 15 });
    const cell3 = table.getCell(0, 0);
    expect(cell3).toEqual({ type: "number", value: 15 });

    table.setCell(0, 0, { type: "date", value: "2020-01-01" });
    const cell4 = table.getCell(0, 0);
    expect(cell4).toEqual({ type: "date", value: "2020-01-01" });

    table.setCell(0, 0, {
      type: "hyperlink",
      value: "https://www.google.com",
    });
    const cell5 = table.getCell(0, 0);
    expect(cell5).toEqual({
      type: "hyperlink",
      value: "https://www.google.com",
    });
  });

  test("should be table", () => {
    const table = new SheetData();
    const cell = table.getCell(3, 4);
    expect(cell).toBeNull();
    expect(table.table).toStrictEqual([]);

    table.setCell(1, 2, { type: "string", value: "hello" });
    const cell2 = table.getCell(1, 2);
    expect(cell2).toEqual({ type: "string", value: "hello" });
    expect(table.table).toStrictEqual([
      [],
      [null, null, { type: "string", value: "hello" }],
    ]);
  });

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
