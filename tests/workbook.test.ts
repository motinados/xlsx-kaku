import { Workbook } from "../src/workbook";

describe("workbook", () => {
  test("The same name will cause an error", async () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");
    expect(() => wb.addWorksheet("Sheet1")).toThrow();
  });

  test("invalid worksheet names cause an error", () => {
    const wb = new Workbook();
    const tooLongName = "a".repeat(32);

    expect(() => wb.addWorksheet("")).toThrow("must not be empty");
    expect(() => wb.addWorksheet(tooLongName)).toThrow(
      "must not be longer than 31 characters"
    );
    expect(() => wb.addWorksheet("Sheet/1")).toThrow("must not contain");
  });

  test("getWorksheet should return the correct worksheet", () => {
    const wb = new Workbook();
    const sheetName = "Sheet1";
    const worksheet = wb.addWorksheet(sheetName);

    const result = wb.getWorksheet(sheetName);

    expect(result).toEqual(worksheet);
  });

  test("getWorksheet should return undefined if worksheet is not found", () => {
    const wb = new Workbook();

    wb.addWorksheet("Sheet2");
    const result = wb.getWorksheet("Sheet1");

    expect(result).toBeUndefined();
  });

  test("generateXlsx should return Uint8Array", async () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");
    const xlsx = await wb.generateXlsx();
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("generateXlsxSync should return Uint8Array", () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");
    const xlsx = wb.generateXlsxSync();
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });
});
