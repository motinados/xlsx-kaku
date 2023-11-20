import { Rows } from "../src/rows";

describe("Rows", () => {
  test("should be able to create a row", () => {
    const rows = new Rows();
    expect(rows).toBeInstanceOf(Rows);
  });

  test("should be able to get a row", () => {
    const rows = new Rows();
    const row = rows.getRow(0);
    expect(row).toEqual({ cells: [] });
  });

  test("should be able to get a row by index", () => {
    const rows = new Rows();
    expect(rows.length).toBe(0);
    const row = rows.getRow(2);
    expect(rows.length).toBe(3);
    expect(row).toEqual({ cells: [] });
  });
});
