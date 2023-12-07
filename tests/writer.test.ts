import { Cell, NullableCell } from "../src/sheetData";
import { SharedStrings } from "../src/sharedStrings";
import {
  cellToString,
  convertIsoStringToSerialValue,
  findFirstNonNullCell,
  findLastNonNullCell,
  getDimension,
  getSpans,
  getSpansFromTable,
  rowToString,
  tableToString,
} from "../src/writer";
import { CellXfs } from "../src/cellXfs";
import { Fonts } from "../src/fonts";
import { Fills } from "../src/fills";
import { Borders } from "../src/borders";
import { NumberFormats } from "../src/numberFormats";
import { CellStyleXfs } from "../src/cellStyleXfs";
import { CellStyles } from "../src/cellStyles";
import { Hyperlinks } from "../src/hyperlinks";
import { WorksheetRels } from "../src/worksheetRels";

describe("Writer", () => {
  test("findFirstNonNullCell", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "name" },
      { type: "string", value: "age" },
    ];
    const { firstNonNullCell, index } = findFirstNonNullCell(row);
    expect(firstNonNullCell).toEqual({ type: "string", value: "name" });
    expect(index).toBe(2);
  });

  test("findLastNonNullCell", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "name" },
      { type: "string", value: "age" },
    ];
    const { lastNonNullCell, index } = findLastNonNullCell(row);
    expect(lastNonNullCell).toEqual({ type: "string", value: "age" });
    expect(index).toBe(3);
  });

  test("getSpans", () => {
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "name" },
      { type: "string", value: "age" },
    ];
    const spans = getSpans(row)!;
    expect(spans.startNumber).toBe(3);
    expect(spans.endNumber).toBe(4);
  });

  test("getSpansFromTable", () => {
    const table: NullableCell[][] = [
      [],
      [null, null, { type: "string", value: "name" }],
      [
        { type: "string", value: "age" },
        { type: "string", value: "age" },
        null,
        null,
      ],
    ];
    const spans = getSpansFromTable(table)!;
    expect(spans.startNumber).toBe(1);
    expect(spans.endNumber).toBe(3);
  });

  test("getDimensions", () => {
    const table: NullableCell[][] = [
      [],
      [null, null, { type: "string", value: "name" }],
      [
        { type: "string", value: "age" },
        { type: "string", value: "age" },
        null,
        null,
      ],
    ];
    const { start, end } = getDimension(table);
    expect(start).toBe("A2");
    expect(end).toBe("C3");
  });

  test("convertISOstringToSerialValue", () => {
    expect(convertIsoStringToSerialValue("2020-01-01T00:00:00.000Z")).toBe(
      43831
    );
    expect(convertIsoStringToSerialValue("2009-12-31T00:00:00.000Z")).toBe(
      40178
    );
    expect(convertIsoStringToSerialValue("2020-11-11T12:00:00.000Z")).toBe(
      44146.5
    );
    expect(convertIsoStringToSerialValue("2023-11-26T15:30:00.000Z")).toBe(
      45256.645833333336
    );
  });

  test("cellToString for number", () => {
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const cell: NonNullable<NullableCell> = {
      type: "number",
      value: 15,
    };
    const result = cellToString(cell, 2, 0, styleMappers);
    expect(result).toBe(`<c r="C1"><v>15</v></c>`);
  });

  test("cellToString for string", () => {
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const cell: NonNullable<NullableCell> = {
      type: "string",
      value: "hello",
    };
    const result = cellToString(cell, 2, 0, styleMappers);
    expect(result).toBe(`<c r="C1" t="s"><v>0</v></c>`);
    expect(styleMappers.sharedStrings.count).toBe(1);
    expect(styleMappers.sharedStrings.uniqueCount).toBe(1);
  });

  test("cellToString for date", () => {
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const cell: Cell = {
      type: "date",
      value: "2020-01-01T00:00:00.000Z",
    };
    const result = cellToString(cell, 2, 0, styleMappers);
    expect(result).toBe(`<c r="C1" s="1"><v>43831</v></c>`);
  });

  test("rowToString for number", () => {
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const row: NullableCell[] = [
      null,
      null,
      { type: "number", value: 15 },
      { type: "number", value: 23 },
    ];
    const result = rowToString(row, 0, 3, 4, styleMappers);
    expect(result).toBe(
      `<row r="1" spans="3:4"><c r="C1"><v>15</v></c><c r="D1"><v>23</v></c></row>`
    );
  });

  test("rowToString for string", () => {
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const row: NullableCell[] = [
      null,
      null,
      { type: "string", value: "hello" },
      { type: "string", value: "world" },
      { type: "string", value: "hello" },
    ];
    const result = rowToString(row, 0, 3, 5, styleMappers);
    expect(result).toBe(
      `<row r="1" spans="3:5"><c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c><c r="E1" t="s"><v>0</v></c></row>`
    );
    expect(styleMappers.sharedStrings.count).toBe(3);
    expect(styleMappers.sharedStrings.uniqueCount).toBe(2);
  });

  test("tableToString for number", () => {
    const table: NullableCell[][] = [
      [],
      [null, null, { type: "number", value: 1 }, { type: "number", value: 2 }],
      [{ type: "number", value: 3 }, { type: "number", value: 4 }, null, null],
    ];
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const result = tableToString(table, styleMappers);
    expect(result.sheetDataXml).toBe(
      `<sheetData><row r="2" spans="1:4"><c r="C2"><v>1</v></c><c r="D2"><v>2</v></c></row><row r="3" spans="1:4"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row></sheetData>`
    );
    expect(result.sharedStringsXml).toBeNull();
  });

  test("tableToString for string", () => {
    const table: NullableCell[][] = [
      [],
      [null, null, { type: "string", value: "hello" }],
      [
        { type: "string", value: "world" },
        { type: "string", value: "world" },
        null,
        null,
      ],
    ];
    const styleMappers = {
      fills: new Fills(),
      fonts: new Fonts(),
      borders: new Borders(),
      numberFormats: new NumberFormats(),
      sharedStrings: new SharedStrings(),
      cellStyleXfs: new CellStyleXfs(),
      cellXfs: new CellXfs(),
      cellStyles: new CellStyles(),
      hyperlinks: new Hyperlinks(),
      worksheetRels: new WorksheetRels(),
    };
    const result = tableToString(table, styleMappers);
    expect(result.sheetDataXml).toBe(
      `<sheetData><row r="2" spans="1:3"><c r="C2" t="s"><v>0</v></c></row><row r="3" spans="1:3"><c r="A3" t="s"><v>1</v></c><c r="B3" t="s"><v>1</v></c></row></sheetData>`
    );
    expect(result.sharedStringsXml).toBe(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2"><si><t>hello</t></si><si><t>world</t></si></sst>`
    );
  });
});
