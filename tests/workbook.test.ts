import { Workbook } from "../src/workbook";

describe("workbook", () => {
  test("The same name will cause an error", async () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");
    expect(() => wb.addWorksheet("Sheet1")).toThrow();
  });

  test("worksheet name with 31 characters is allowed", () => {
    const wb = new Workbook();
    const name = "a".repeat(31);

    const ws = wb.addWorksheet(name);

    expect(ws.name).toBe(name);
  });

  test("worksheet name with 32 characters causes an error", () => {
    const wb = new Workbook();
    const name = "a".repeat(32);

    expect(() => wb.addWorksheet(name)).toThrow(
      "must not be longer than 31 characters"
    );
  });

  test("empty worksheet name causes an error", () => {
    const wb = new Workbook();

    expect(() => wb.addWorksheet("")).toThrow("must not be empty");
  });

  test.each(["/", "\\", "?", "*", ":", "[", "]"])(
    "invalid worksheet name containing %s causes an error",
    (invalidChar) => {
      const wb = new Workbook();

      expect(() => wb.addWorksheet(`Sheet${invalidChar}1`)).toThrow(
        "must not contain"
      );
    }
  );

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
