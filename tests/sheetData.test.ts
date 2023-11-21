import { SheetData } from "../src/sheetData";

describe("SheetData", () => {
  test("should be able to create a sheetData", () => {
    const sheetData = new SheetData();
    expect(sheetData).toBeInstanceOf(SheetData);
  });

  test("should be able to get a cell", () => {
    const sheetData = new SheetData();
    const cell = sheetData.getCell(0, 0);
    expect(cell).not.toBeUndefined();
  });

  test("should be able to get a cell by index", () => {
    const sheetData = new SheetData();
    expect(sheetData.rowsLength).toBe(0);
    const cell = sheetData.getCell(2, 2);
    expect(sheetData.rowsLength).toBe(3);
    expect(cell.value).toEqual("");
  });

  test("should be able to get a cell by index", () => {
    const sheetData = new SheetData();

    const cell = sheetData.getCell(0, 0);
    expect(cell.type).toEqual("string");
    expect(cell.value).toEqual("");

    cell.value = 15;

    expect(sheetData.getCell(0, 0).type).toEqual("number");
    expect(sheetData.getCell(0, 0).value).toEqual(15);
  });
});
