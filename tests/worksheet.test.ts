import { Worksheet } from "../src";

describe("Worksheet", () => {
  test("should be able to create a sheet", () => {
    const sheet = new Worksheet();
    expect(sheet).toBeInstanceOf(Worksheet);
  });

  test("should be able to add a row", () => {
    const sheet = new Worksheet();
    sheet.addRow({ cells: [{ value: "hello" }] });
    expect(sheet.getRows().length).toBe(1);
  });
});
