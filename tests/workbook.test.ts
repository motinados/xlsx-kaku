import { Workbook } from "../src/workbook";

describe("Sheet", () => {
  test("should be able to create a workbook", () => {
    const workbook = new Workbook();
    expect(workbook).toBeInstanceOf(Workbook);
  });
});
