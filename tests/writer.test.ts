import { Cell, RowData, SheetData } from "../src/sheetData";
import { SharedStrings } from "../src/sharedStrings";
import {
  cellToString,
  convertIsoStringToSerialValue,
  findFirstNonNullCell,
  findLastNonNullCell,
  getDimension,
  getSpans,
  getSpansFromSheetData,
  makeColsXml,
  makeMergeCellsXml,
  makeSheetDataXml,
  makeSheetViewsXml,
  rowToString,
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
import { Col, FreezePane, MergeCell } from "../src/worksheet";

describe("Writer", () => {
  test("findFirstNonNullCell", () => {
    const row: RowData = [
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
    const row: RowData = [
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
    const row: RowData = [
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
    const sheetData: SheetData = [
      [],
      [null, null, { type: "string", value: "name" }],
      [
        { type: "string", value: "age" },
        { type: "string", value: "age" },
        null,
        null,
      ],
    ];
    const spans = getSpansFromSheetData(sheetData)!;
    expect(spans.startNumber).toBe(1);
    expect(spans.endNumber).toBe(3);
  });

  test("getDimensions", () => {
    const sheetData: SheetData = [
      [],
      [null, null, { type: "string", value: "name" }],
      [
        { type: "string", value: "age" },
        { type: "string", value: "age" },
        null,
        null,
      ],
    ];
    const { start, end } = getDimension(sheetData);
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
    const cell: Cell = {
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
    const cell: Cell = {
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

  test("cellToString for Hyperlink", () => {
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
      type: "hyperlink",
      value: "https://www.google.com",
    };
    const result = cellToString(cell, 2, 0, styleMappers);
    expect(result).toBe(`<c r="C1" s="1" t="s"><v>0</v></c>`);

    const worksheetRels = styleMappers.worksheetRels.getWorksheetRels();
    expect(worksheetRels.length).toBe(1);
    const rels = worksheetRels[0];
    if (rels === undefined) {
      throw new Error("rels is undefined");
    }

    expect(rels.target).toBe("https://www.google.com");
    expect(rels.targetMode).toBe("External");
    const rid = rels.id;

    const hyperlinks = styleMappers.hyperlinks.getHyperlinks();
    expect(hyperlinks.length).toBe(1);
    const hyperlink = hyperlinks[0];
    if (hyperlink === undefined) {
      throw new Error("hyperlink is undefined");
    }

    expect(hyperlink).not.toBeUndefined();
    expect(hyperlink.ref).toBe("C1");
    expect(hyperlink.rid).toBe(rid);
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
    const row: RowData = [
      null,
      null,
      { type: "number", value: 15 },
      { type: "number", value: 23 },
    ];
    const result = rowToString(row, 0, null, 3, 4, styleMappers);
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
    const row: RowData = [
      null,
      null,
      { type: "string", value: "hello" },
      { type: "string", value: "world" },
      { type: "string", value: "hello" },
    ];
    const result = rowToString(row, 0, null, 3, 5, styleMappers);
    expect(result).toBe(
      `<row r="1" spans="3:5"><c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c><c r="E1" t="s"><v>0</v></c></row>`
    );
    expect(styleMappers.sharedStrings.count).toBe(3);
    expect(styleMappers.sharedStrings.uniqueCount).toBe(2);
  });

  test("rowToString with height", () => {
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
    const row: RowData = [{ type: "number", value: 10 }];
    const result = rowToString(row, 0, 30, 1, 1, styleMappers);
    expect(result).toBe(
      `<row r="1" spans="1:1" ht="30" customHeight="1"><c r="A1"><v>10</v></c></row>`
    );
  });

  test("tableToString for number", () => {
    const sheetData: SheetData = [
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
    const sheetDataXml = makeSheetDataXml(sheetData, [], styleMappers);
    expect(sheetDataXml).toBe(
      `<sheetData><row r="2" spans="1:4"><c r="C2"><v>1</v></c><c r="D2"><v>2</v></c></row><row r="3" spans="1:4"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row></sheetData>`
    );
  });

  test("tableToString for string", () => {
    const sheetData: SheetData = [
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
    const sheetDataXml = makeSheetDataXml(sheetData, [], styleMappers);
    expect(sheetDataXml).toBe(
      `<sheetData><row r="2" spans="1:3"><c r="C2" t="s"><v>0</v></c></row><row r="3" spans="1:3"><c r="A3" t="s"><v>1</v></c><c r="B3" t="s"><v>1</v></c></row></sheetData>`
    );
  });

  test("makeColsXml", () => {
    const cols: Col[] = [
      { min: 1, max: 1, width: 10 },
      { min: 2, max: 2, width: 75 },
      { min: 3, max: 6, width: 25 },
    ];

    expect(makeColsXml(cols)).toBe(
      `<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="75" customWidth="1"/><col min="3" max="6" width="25" customWidth="1"/></cols>`
    );
  });

  test("makeMergeCellsXml", () => {
    const mergeCells: MergeCell[] = [
      { ref: "A1:B2" },
      { ref: "C3:D4" },
      { ref: "E5:F6" },
    ];

    expect(makeMergeCellsXml(mergeCells)).toBe(
      `<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="C3:D4"/><mergeCell ref="E5:F6"/></mergeCells>`
    );
  });

  test("mekeSheetViewsXml", () => {
    const dimension = { start: "A1", end: "B2" };
    expect(makeSheetViewsXml(dimension, null)).toBe(
      `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });

  test("mekeSheetViewsXml with frozen column", () => {
    const dimension = { start: "A1", end: "B2" };
    const freezePane: FreezePane = { type: "column", split: 1 };
    expect(makeSheetViewsXml(dimension, freezePane)).toBe(
      `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });

  test("mekeSheetViewsXml with frozen row", () => {
    const dimension = { start: "A1", end: "B2" };
    const freezePane: FreezePane = { type: "row", split: 1 };
    expect(makeSheetViewsXml(dimension, freezePane)).toBe(
      `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });
});
