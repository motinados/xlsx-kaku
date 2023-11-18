import { Worksheet } from "../src/index";
import { Workbook } from "../src/workbook";

describe("Sheet", () => {
  test("should be able to create a workbook", () => {
    const workbook = new Workbook();
    expect(workbook).toBeInstanceOf(Workbook);
  });

  test("should be able to create a sheet", () => {
    const sheet = new Worksheet();
    expect(sheet).toBeInstanceOf(Worksheet);
  });
});
