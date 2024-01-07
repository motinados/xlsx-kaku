import { Workbook } from "../src/workbook";

describe("workbook", () => {
  test("The same name will cause an error", async () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");
    expect(() => wb.addWorksheet("Sheet1")).toThrow();
  });
});
