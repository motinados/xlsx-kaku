import { Worksheet } from "../src/worksheet";

describe("worksheet", () => {
  test("should be able to create a worksheet", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws).toBeInstanceOf(Worksheet);
  });

  test("get name", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws.name).toBe("Sheet1");
  });

  test("set sheetData", () => {
    const ws = new Worksheet("Sheet1");
    ws.sheetData = [[{ type: "string", value: "Hello" }]];
    expect(ws.sheetData).toStrictEqual([[{ type: "string", value: "Hello" }]]);

    ws.sheetData = [
      [{ type: "string", value: "Hello" }],
      [{ type: "string", value: "World" }],
    ];
    expect(ws.sheetData).toStrictEqual([
      [{ type: "string", value: "Hello" }],
      [{ type: "string", value: "World" }],
    ]);
  });

  test("setCell", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(0, 0, { type: "string", value: "Hello" });
    expect(ws.sheetData).toStrictEqual([[{ type: "string", value: "Hello" }]]);
    ws.setCell(0, 1, { type: "string", value: "World" });
    expect(ws.sheetData).toStrictEqual([
      [
        { type: "string", value: "Hello" },
        { type: "string", value: "World" },
      ],
    ]);
  });

  test("setCell with empty", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(0, 1, { type: "string", value: "Hello" });
    expect(ws.sheetData).toStrictEqual([
      [null, { type: "string", value: "Hello" }],
    ]);
    ws.setCell(3, 0, { type: "string", value: "World" });
    expect(ws.sheetData).toStrictEqual([
      [null, { type: "string", value: "Hello" }],
      [],
      [],
      [{ type: "string", value: "World" }],
    ]);
  });

  test("set sheetData and setCell", () => {
    const ws = new Worksheet("Sheet1");
    ws.sheetData = [[{ type: "string", value: "Hello" }]];
    ws.setCell(0, 1, { type: "string", value: "World" });
    expect(ws.sheetData).toStrictEqual([
      [
        { type: "string", value: "Hello" },
        { type: "string", value: "World" },
      ],
    ]);
  });
});
