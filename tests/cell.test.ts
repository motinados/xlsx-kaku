import { Table } from "../src/cell";

describe("Cell", () => {
  test("should be able to create a cell", () => {
    const table = new Table();
    expect(table).toBeInstanceOf(Table);
  });

  test("should be able to get a cell", () => {
    const table = new Table();
    const cell = table.getCell(0, 0);
    expect(cell).toBeNull();
  });

  test("should be able to set a cell", () => {
    const table = new Table();
    const cell = table.getCell(0, 0);
    expect(cell).toBeNull();

    table.setCellElm(0, 0, { type: "string", value: "hello" });
    const cell2 = table.getCell(0, 0);
    expect(cell2).toEqual({ type: "string", value: "hello" });

    table.setCellElm(0, 0, { type: "number", value: 15 });
    const cell3 = table.getCell(0, 0);
    expect(cell3).toEqual({ type: "number", value: 15 });

    table.setCellElm(0, 0, { type: "date", value: "2020-01-01" });
    const cell4 = table.getCell(0, 0);
    expect(cell4).toEqual({ type: "date", value: "2020-01-01" });
  });

  test("should be table", () => {
    const table = new Table();
    const cell = table.getCell(3, 4);
    expect(cell).toBeNull();
    expect(table.table).toStrictEqual([]);

    table.setCellElm(1, 2, { type: "string", value: "hello" });
    const cell2 = table.getCell(1, 2);
    expect(cell2).toEqual({ type: "string", value: "hello" });
    expect(table.table).toStrictEqual([
      [],
      [null, null, { type: "string", value: "hello" }],
    ]);
  });
});
