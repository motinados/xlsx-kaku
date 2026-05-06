import { ImageStore } from "../src/imageStore";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  Worksheet,
} from "../src/worksheet";

describe("worksheet", () => {
  test("should be able to create a worksheet", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws).toBeInstanceOf(Worksheet);
  });

  test("worksheet name with 31 characters is allowed", () => {
    const name = "a".repeat(31);

    const ws = new Worksheet(name);

    expect(ws.name).toBe(name);
  });

  test("worksheet name with 32 characters causes an error", () => {
    const name = "a".repeat(32);

    expect(() => new Worksheet(name)).toThrow(
      "must not be longer than 31 characters"
    );
  });

  test("empty worksheet name causes an error", () => {
    expect(() => new Worksheet("")).toThrow("must not be empty");
  });

  test.each(["/", "\\", "?", "*", ":", "[", "]"])(
    "invalid worksheet name containing %s causes an error",
    (invalidChar) => {
      expect(() => new Worksheet(`Sheet${invalidChar}1`)).toThrow(
        "must not contain"
      );
    }
  );

  test("get name", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws.name).toBe("Sheet1");
  });

  test("get opts", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws.opts).toStrictEqual({
      defaultColWidth: DEFAULT_COL_WIDTH,
      defaultRowHeight: DEFAULT_ROW_HEIGHT,
    });

    const ws2 = new Worksheet("Sheet2", new ImageStore(), {
      defaultColWidth: 10,
    });
    expect(ws2.opts).toStrictEqual({
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

  test("setCell overwrites existing cell", () => {
    const ws = new Worksheet("Sheet1");

    ws.setCell(0, 0, { type: "string", value: "Hello" });
    ws.setCell(0, 0, { type: "string", value: "World" });

    expect(ws.sheetData).toStrictEqual([[{ type: "string", value: "World" }]]);
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

  test("setCell fills missing columns with null in existing row", () => {
    const ws = new Worksheet("Sheet1");

    ws.setCell(0, 0, { type: "string", value: "A" });
    ws.setCell(0, 2, { type: "string", value: "C" });

    expect(ws.sheetData).toStrictEqual([
      [{ type: "string", value: "A" }, null, { type: "string", value: "C" }],
    ]);
  });

  test("getCell returns null for empty sheet", () => {
    const ws = new Worksheet("Sheet1");
    expect(ws.getCell(0, 0)).toBeNull();
  });

  test("getCell returns null for explicitly empty cell", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(0, 1, { type: "string", value: "Hello" });
    expect(ws.getCell(0, 0)).toBeNull();
  });

  test("getCell returns null if row does not exist", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    expect(ws.getCell(0, 0)).toBeNull();
    expect(ws.getCell(0, 1)).toBeNull();
  });

  test("getCell returns null if col does not exist", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(0, 0, { type: "string", value: "Hello" });
    expect(ws.getCell(0, 1)).toBeNull();
  });

  test("getCell returns set value", () => {
    const ws = new Worksheet("Sheet1");
    const cell = { type: "string", value: "Hello" } as const;
    ws.setCell(0, 0, cell);
    expect(ws.getCell(0, 0)).toStrictEqual(cell);
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
