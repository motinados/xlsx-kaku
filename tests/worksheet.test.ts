import { DEFAULT_ROW_HEIGHT } from "../src/row";
import { DEFAULT_COL_WIDTH, Worksheet } from "../src/worksheet";

describe("worksheet", () => {
  test("should be able to create a worksheet", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws).toBeInstanceOf(Worksheet);
  });

  test("get name", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws.name).toBe("Sheet1");
  });

  test("get props", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws.props).toStrictEqual({
      defaultColWidth: DEFAULT_COL_WIDTH,
      defaultRowHeight: DEFAULT_ROW_HEIGHT,
    });

    const ws2 = new Worksheet("Sheet2", { defaultColWidth: 10 });
    expect(ws2.props).toStrictEqual({
      defaultColWidth: 10,
      defaultRowHeight: DEFAULT_ROW_HEIGHT,
    });
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

  test("setMergeCell", () => {
    const ws = new Worksheet("Sheet1");
    const mergeCell = { ref: "A1:B2" };
    ws.setMergeCell(mergeCell);

    const expectedSheetData = [
      [{ type: "string", value: "" }, { type: "merged" }],
      [{ type: "merged" }, { type: "merged" }],
    ];

    expect(ws.sheetData).toStrictEqual(expectedSheetData);
    expect(ws.mergeCells).toStrictEqual([mergeCell]);
  });
});
