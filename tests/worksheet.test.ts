import { Worksheet } from "../src";

describe("Worksheet", () => {
  test("should be able to create a sheet", () => {
    const sheet = new Worksheet({ sheetName: "Sheet1" });
    expect(sheet).toBeInstanceOf(Worksheet);
  });

  test("should be able to set a sheet name", () => {
    const sheet = new Worksheet({ sheetName: "Sheet1" });
    expect(sheet.sheetName).toBe("Sheet1");
  });

  test("should be able to change a sheet name", () => {
    const sheet = new Worksheet({ sheetName: "Sheet1" });
    sheet.sheetName = "Sheet2";
    expect(sheet.sheetName).toBe("Sheet2");
  });

  test("should be able to add a row", () => {
    const sheet = new Worksheet({ sheetName: "Sheet1" });
    sheet.addRow({ cells: [] });
    expect(sheet.getRows().length).toBe(1);
  });

  test("should be able to get all rows", () => {
    const sheet = new Worksheet({ sheetName: "Sheet1" });
    sheet.addRow({ cells: [] });
    sheet.addRow({ cells: [] });
    sheet.addRow({ cells: [] });
    expect(sheet.getRows().length).toBe(3);
  });
});
