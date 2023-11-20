import { SheetData } from "../src/sheetData";

describe("SheetData", () => {
  test("should be able to create a sheetData", () => {
    const sheetData = new SheetData();
    expect(sheetData).toBeInstanceOf(SheetData);
  });
  test("should be able to get a cell", () => {
    const sheetData = new SheetData();
    const cell = sheetData.getCell(0, 0);
    expect(cell).toEqual({ value: "" });
  });
  test("should be able to get a cell by index", () => {
    const sheetData = new SheetData();
    expect(sheetData.rowsLength).toBe(0);
    const cell = sheetData.getCell(2, 2);
    expect(sheetData.rowsLength).toBe(3);
    expect(cell).toEqual({ value: "" });
  });
});
