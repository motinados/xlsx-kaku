import { Cell, RowData, SheetData } from "../../src/sheetData";
import { SharedStrings } from "../../src/sharedStrings";
import {
  convertCellToXlsxCell,
  convertIsoStringToSerialValue,
  createXlsxColFromColProps,
  createXlsxRowFromRowProps,
  findFirstNonNullCell,
  findLastNonNullCell,
  getDimension,
  getSpans,
  getSpansFromSheetData,
  makeCellXml,
  makeColsXml,
  makeMergeCellsXml,
  makeSheetDataXml,
  makeSheetViewsXml,
  rowToString,
  makeSheetFormatPrXml,
  groupXlsxCols,
  XlsxCol,
  isEqualsXlsxCol,
  GroupedXlsxCol,
  makeConditionalFormattingXml,
  XlsxConditionalFormatting,
} from "../../src/xml/worksheetXml";
import { CellXfs } from "../../src/cellXfs";
import { Fonts } from "../../src/fonts";
import { Fills } from "../../src/fills";
import { Borders } from "../../src/borders";
import { NumberFormats } from "../../src/numberFormats";
import { CellStyleXfs } from "../../src/cellStyleXfs";
import { CellStyles } from "../../src/cellStyles";
import { Hyperlinks } from "../../src/hyperlinks";
import { WorksheetRels } from "../../src/worksheetRels";
import {
  ColProps,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  FreezePane,
  MergeCell,
} from "../../src/worksheet";

describe("Writer", () => {
  function getStyleMappers() {
    return {
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
  }

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
    expect(spans.spanStartNumber).toBe(1);
    expect(spans.spanEndNumber).toBe(3);
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
    const { spanStartNumber, spanEndNumber } =
      getSpansFromSheetData(sheetData)!;
    const { start, end } = getDimension(
      sheetData,
      spanStartNumber,
      spanEndNumber
    );
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
    const styleMappers = getStyleMappers();
    const cell: Cell = {
      type: "number",
      value: 15,
    };
    const result = makeCellXml(
      convertCellToXlsxCell(cell, 2, 0, styleMappers, new Map(), undefined)
    );
    expect(result).toBe(`<c r="C1"><v>15</v></c>`);
  });

  test("cellToString for string", () => {
    const styleMappers = getStyleMappers();
    const cell: Cell = {
      type: "string",
      value: "hello",
    };
    const result = makeCellXml(
      convertCellToXlsxCell(cell, 2, 0, styleMappers, new Map(), undefined)
    );
    expect(result).toBe(`<c r="C1" t="s"><v>0</v></c>`);
    expect(styleMappers.sharedStrings.count).toBe(1);
    expect(styleMappers.sharedStrings.uniqueCount).toBe(1);
  });

  test("cellToString for date", () => {
    const styleMappers = getStyleMappers();
    const cell: Cell = {
      type: "date",
      value: "2020-01-01T00:00:00.000Z",
    };
    const result = makeCellXml(
      convertCellToXlsxCell(cell, 2, 0, styleMappers, new Map(), undefined)
    );
    expect(result).toBe(`<c r="C1" s="1"><v>43831</v></c>`);
  });

  test("cellToString for Hyperlink", () => {
    const styleMappers = getStyleMappers();
    const cell: Cell = {
      type: "hyperlink",
      text: "https://www.google.com",
      value: "https://www.google.com",
      linkType: "external",
    };
    const result = makeCellXml(
      convertCellToXlsxCell(cell, 2, 0, styleMappers, new Map(), undefined)
    );
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
    if (hyperlink.linkType !== "external") {
      throw new Error("hyperlink.linkType is not external");
    }

    expect(hyperlink).not.toBeUndefined();
    expect(hyperlink.ref).toBe("C1");
    expect(hyperlink.rid).toBe(rid);
  });

  test("rowToString for number", () => {
    const styleMappers = getStyleMappers();
    const row: RowData = [
      null,
      null,
      { type: "number", value: 15 },
      { type: "number", value: 23 },
    ];
    const result = rowToString(
      row,
      0,
      3,
      4,
      styleMappers,
      new Map(),
      new Map()
    );
    expect(result).toBe(
      `<row r="1" spans="3:4"><c r="C1"><v>15</v></c><c r="D1"><v>23</v></c></row>`
    );
  });

  test("rowToString for string", () => {
    const styleMappers = getStyleMappers();
    const row: RowData = [
      null,
      null,
      { type: "string", value: "hello" },
      { type: "string", value: "world" },
      { type: "string", value: "hello" },
    ];

    const result = rowToString(
      row,
      0,
      3,
      5,
      styleMappers,
      new Map(),
      new Map()
    );
    expect(result).toBe(
      `<row r="1" spans="3:5"><c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c><c r="E1" t="s"><v>0</v></c></row>`
    );
    expect(styleMappers.sharedStrings.count).toBe(3);
    expect(styleMappers.sharedStrings.uniqueCount).toBe(2);
  });

  test("rowToString with height", () => {
    const styleMappers = getStyleMappers();
    const row: RowData = [{ type: "number", value: 10 }];
    const xlsxRow = createXlsxRowFromRowProps(
      { index: 0, height: 30 },
      styleMappers
    );
    const result = rowToString(
      row,
      0,
      1,
      1,
      styleMappers,
      new Map(),
      new Map([[0, xlsxRow]])
    );
    expect(result).toBe(
      `<row r="1" spans="1:1" ht="30" customHeight="1"><c r="A1"><v>10</v></c></row>`
    );
  });

  test("rowToString with style", () => {
    const styleMappers = getStyleMappers();
    const row: RowData = [{ type: "number", value: 10 }];
    const xlsxRow = createXlsxRowFromRowProps(
      { index: 0, style: { alignment: { horizontal: "center" } } },
      styleMappers
    );
    const result = rowToString(
      row,
      0,
      1,
      1,
      styleMappers,
      new Map(),
      new Map([[0, xlsxRow]])
    );
    expect(result).toBe(
      `<row r="1" spans="1:1" s="1" customFormat="1"><c r="A1" s="1"><v>10</v></c></row>`
    );
  });

  test("rowToString with style and height", () => {
    const styleMappers = getStyleMappers();
    const row: RowData = [{ type: "number", value: 10 }];
    const xlsxRow = createXlsxRowFromRowProps(
      {
        index: 0,
        height: 30,
        style: { alignment: { horizontal: "center" } },
      },
      styleMappers
    );
    const result = rowToString(
      row,
      0,
      1,
      1,
      styleMappers,
      new Map(),
      new Map([[0, xlsxRow]])
    );
    expect(result).toBe(
      `<row r="1" spans="1:1" s="1" customFormat="1" ht="30" customHeight="1"><c r="A1" s="1"><v>10</v></c></row>`
    );
  });

  test("tableToString for number", () => {
    const sheetData: SheetData = [
      [],
      [null, null, { type: "number", value: 1 }, { type: "number", value: 2 }],
      [{ type: "number", value: 3 }, { type: "number", value: 4 }, null, null],
    ];
    const { spanStartNumber, spanEndNumber } = getSpansFromSheetData(sheetData);
    const styleMappers = getStyleMappers();
    const sheetDataXml = makeSheetDataXml(
      sheetData,
      spanStartNumber,
      spanEndNumber,
      styleMappers,
      new Map(),
      new Map()
    );
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
    const { spanStartNumber, spanEndNumber } = getSpansFromSheetData(sheetData);
    const styleMappers = getStyleMappers();
    const sheetDataXml = makeSheetDataXml(
      sheetData,
      spanStartNumber,
      spanEndNumber,
      styleMappers,
      new Map(),
      new Map()
    );
    expect(sheetDataXml).toBe(
      `<sheetData><row r="2" spans="1:3"><c r="C2" t="s"><v>0</v></c></row><row r="3" spans="1:3"><c r="A3" t="s"><v>1</v></c><c r="B3" t="s"><v>1</v></c></row></sheetData>`
    );
  });

  test("makeColsXml for width", () => {
    const styleMappers = getStyleMappers();
    const cols: ColProps[] = [
      { index: 0, width: 10 },
      { index: 1, width: 75 },
      { index: 2, width: 25 },
      { index: 3, width: 25 },
      { index: 4, width: 25 },
      { index: 5, width: 25 },
    ];

    const xlsxCols = new Map<number, XlsxCol>();
    cols.forEach((col) => {
      xlsxCols.set(
        col.index,
        createXlsxColFromColProps(col, styleMappers, DEFAULT_COL_WIDTH)
      );
    });
    const groupedXlsxCols = groupXlsxCols(xlsxCols);
    expect(makeColsXml(groupedXlsxCols, DEFAULT_COL_WIDTH)).toBe(
      `<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="75" customWidth="1"/><col min="3" max="6" width="25" customWidth="1"/></cols>`
    );
  });

  test("makeColsXml for default width", () => {
    const styleMappers = getStyleMappers();
    const cols: ColProps[] = [{ index: 0, width: DEFAULT_COL_WIDTH }];

    const xlsxCols = new Map<number, XlsxCol>();
    cols.forEach((col) => {
      xlsxCols.set(
        col.index,
        createXlsxColFromColProps(col, styleMappers, DEFAULT_COL_WIDTH)
      );
    });
    const groupedXlsxCols = groupXlsxCols(xlsxCols);

    // TODO: Is this necessary?
    expect(makeColsXml(groupedXlsxCols, DEFAULT_COL_WIDTH)).toBe(
      `<cols><col min="1" max="1" width="${DEFAULT_COL_WIDTH}"/></cols>`
    );
  });

  test("makeColsXml for style", () => {
    const styleMappers = getStyleMappers();
    const cols: ColProps[] = [
      {
        index: 0,
        style: { alignment: { horizontal: "center" } },
      },
      {
        index: 1,
        width: 25,
        style: {
          fill: { patternType: "solid", fgColor: "FFFF0000" },
        },
      },
      {
        index: 2,
        width: 25,
        style: {
          fill: { patternType: "solid", fgColor: "FFFF0000" },
        },
      },
    ];

    const xlsxCols = new Map<number, XlsxCol>();
    cols.forEach((col) => {
      xlsxCols.set(
        col.index,
        createXlsxColFromColProps(col, styleMappers, DEFAULT_COL_WIDTH)
      );
    });
    const groupedXlsxCols = groupXlsxCols(xlsxCols);
    expect(makeColsXml(groupedXlsxCols, DEFAULT_COL_WIDTH)).toBe(
      `<cols><col min="1" max="1" width="${DEFAULT_COL_WIDTH}" style="1"/><col min="2" max="3" width="25" customWidth="1" style="2"/></cols>`
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
    expect(makeSheetViewsXml(true, dimension, null)).toBe(
      `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });

  test("mekeSheetViewsXml with tabSelected false", () => {
    const dimension = { start: "A1", end: "B2" };
    expect(makeSheetViewsXml(false, dimension, null)).toBe(
      `<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });

  test("mekeSheetViewsXml with frozen column", () => {
    const dimension = { start: "A1", end: "B2" };
    const freezePane: FreezePane = { target: "column", split: 1 };
    expect(makeSheetViewsXml(true, dimension, freezePane)).toBe(
      `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });

  test("mekeSheetViewsXml with frozen row", () => {
    const dimension = { start: "A1", end: "B2" };
    const freezePane: FreezePane = { target: "row", split: 1 };
    expect(makeSheetViewsXml(true, dimension, freezePane)).toBe(
      `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
    );
  });

  test("makeSheetFormatPrXml", () => {
    expect(makeSheetFormatPrXml(10, 20)).toBe(
      `<sheetFormatPr defaultRowHeight="10" defaultColWidth="20"/>`
    );

    expect(makeSheetFormatPrXml(10, DEFAULT_COL_WIDTH)).toBe(
      `<sheetFormatPr defaultRowHeight="10"/>`
    );

    expect(makeSheetFormatPrXml(DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH)).toBe(
      `<sheetFormatPr defaultRowHeight="${DEFAULT_ROW_HEIGHT}"/>`
    );
  });

  test("isEqualsXlsxCol", () => {
    const col0: XlsxCol = {
      index: 0,
      width: DEFAULT_COL_WIDTH,
      customWidth: false,
      cellXfId: 1,
    };
    const col1: XlsxCol = {
      index: 1,
      width: 25,
      customWidth: true,
      cellXfId: 2,
    };
    const col2: XlsxCol = {
      index: 1,
      width: 25,
      customWidth: true,
      cellXfId: 2,
    };
    expect(isEqualsXlsxCol(col0, col1)).toBe(false);
    expect(isEqualsXlsxCol(col1, col2)).toBe(true);
  });

  test("groupXlsxCol 1", () => {
    // a, b
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 0,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 1,
        endIndex: 1,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("groupXlsxCol 2", () => {
    // a, b, b
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c3: XlsxCol = {
      index: 2,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
        [2, c3],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 0,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 1,
        endIndex: 2,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("groupXlsxCol 3", () => {
    // a, a, b
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c3: XlsxCol = {
      index: 2,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
        [2, c3],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 1,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 2,
        endIndex: 2,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("groupXlsxCol 4", () => {
    // a, a, a, b
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c3: XlsxCol = {
      index: 2,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c4: XlsxCol = {
      index: 3,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
        [2, c3],
        [3, c4],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 2,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 3,
        endIndex: 3,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("groupXlsxCol 5", () => {
    // a, b, b, b
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c3: XlsxCol = {
      index: 2,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c4: XlsxCol = {
      index: 3,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
        [2, c3],
        [3, c4],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 0,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 1,
        endIndex: 3,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("groupXlsxCol 6", () => {
    // a, a, a, b, a
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c3: XlsxCol = {
      index: 2,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c4: XlsxCol = {
      index: 3,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c5: XlsxCol = {
      index: 4,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
        [2, c3],
        [3, c4],
        [4, c5],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 2,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 3,
        endIndex: 3,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
      {
        startIndex: 4,
        endIndex: 4,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("groupXlsxCol 7", () => {
    // a, b, b, b, a
    const c1: XlsxCol = {
      index: 0,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const c2: XlsxCol = {
      index: 1,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c3: XlsxCol = {
      index: 2,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c4: XlsxCol = {
      index: 3,
      width: 20,
      cellXfId: 1,
      customWidth: true,
    };
    const c5: XlsxCol = {
      index: 4,
      width: 10,
      cellXfId: 0,
      customWidth: true,
    };
    const actual = groupXlsxCols(
      new Map([
        [0, c1],
        [1, c2],
        [2, c3],
        [3, c4],
        [4, c5],
      ])
    );
    const expected: GroupedXlsxCol[] = [
      {
        startIndex: 0,
        endIndex: 0,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
      {
        startIndex: 1,
        endIndex: 3,
        width: 20,
        cellXfId: 1,
        customWidth: true,
      },
      {
        startIndex: 4,
        endIndex: 4,
        width: 10,
        cellXfId: 0,
        customWidth: true,
      },
    ];
    expect(actual).toEqual(expected);
  });

  test("makeConditionalFormattingXml", () => {
    const conditionalFormattings: XlsxConditionalFormatting[] = [
      {
        sqref: "A1:A10",
        bottom: false,
        dxfId: 0,
        priority: 1,
        type: "top10",
        rank: 10,
        percent: true,
      },
    ];
    const actual = makeConditionalFormattingXml(conditionalFormattings);
    const expected =
      `<conditionalFormatting sqref="A1:A10">` +
      `<cfRule type="top10" dxfId="0" priority="1" percent="1" rank="10"/>` +
      `</conditionalFormatting>`;
    expect(actual).toBe(expected);

    conditionalFormattings.push({
      sqref: "B1:B10",
      bottom: true,
      dxfId: 1,
      priority: 1,
      type: "top10",
      rank: 10,
      percent: false,
    });

    const acutual2 = makeConditionalFormattingXml(conditionalFormattings);
    const expected2 =
      `<conditionalFormatting sqref="A1:A10">` +
      `<cfRule type="top10" dxfId="0" priority="1" percent="1" rank="10"/>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="B1:B10">` +
      `<cfRule type="top10" dxfId="1" priority="1" bottom="1" rank="10"/>` +
      `</conditionalFormatting>`;

    expect(acutual2).toBe(expected2);
  });
});
