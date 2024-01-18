import { v4 as uuidv4 } from "uuid";
import { Cell, CellStyle, RowData, SheetData } from "./sheetData";
import { SharedStrings } from "./sharedStrings";
import { makeThemeXml } from "./theme";
import { Fills } from "./fills";
import { Alignment, CellXf, CellXfs } from "./cellXfs";
import { Fonts } from "./fonts";
import { Borders } from "./borders";
import { NumberFormats } from "./numberFormats";
import { CellStyles } from "./cellStyles";
import { CellStyleXfs } from "./cellStyleXfs";
import { Hyperlinks } from "./hyperlinks";
import { WorksheetRels } from "./worksheetRels";
import { FreezePane, MergeCell, Worksheet } from "./worksheet";
import { convColIndexToColName, isInRange } from "./utils";
import { CombinedCol, DEFAULT_COL_WIDTH, combineColProps } from "./col";
import { strToU8, zipSync } from "fflate";
import { CombinedRow, DEFAULT_ROW_HEIGHT, combineRowProps } from "./row";

type StyleMappers = {
  fills: Fills;
  fonts: Fonts;
  borders: Borders;
  numberFormats: NumberFormats;
  sharedStrings: SharedStrings;
  cellStyleXfs: CellStyleXfs;
  cellXfs: CellXfs;
  cellStyles: CellStyles;
  hyperlinks: Hyperlinks;
  worksheetRels: WorksheetRels;
};

type XlsxCol = {
  /** e.g. column A is 1 */
  min: number;
  /** e.g. column A is 1 */
  max: number;
  width: number;
  customWidth: boolean;
  cellXfId: number | null;
};

type XlsxRow = {
  /** e.g. rows[0] is 0 */
  index: number;
  height: number;
  customHeight: boolean;
  cellXfId: number | null;
};

type XlsxCellStyle = {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: Alignment;
};

type XlsxCell =
  | {
      type: "number";
      colName: string;
      rowNumber: number;
      value: number;
      cellXfId: number | null;
    }
  | {
      type: "string";
      colName: string;
      rowNumber: number;
      value: string;
      sharedStringId: number;
      cellXfId: number | null;
    }
  | {
      type: "date";
      colName: string;
      rowNumber: number;
      value: string;
      cellXfId: number | null;
    }
  | {
      type: "hyperlink";
      colName: string;
      rowNumber: number;
      value: string;
      sharedStringId: number;
      cellXfId: number | null;
    }
  | {
      type: "boolean";
      colName: string;
      rowNumber: number;
      value: boolean;
      cellXfId: number | null;
    }
  | {
      type: "formula";
      colName: string;
      rowNumber: number;
      value: string;
      cellXfId: number | null;
    }
  | {
      type: "merged";
      colName: string;
      rowNumber: number;
      cellXfId: number | null;
    };

export function genXlsx(worksheets: Worksheet[]) {
  const files = generateXMLs(worksheets);
  const zipped = compressXMLs(files);
  return zipped;
}

function compressXMLs(files: { filename: string; content: string }[]) {
  const data: { [key: string]: Uint8Array } = {};

  for (const file of files) {
    data[file.filename] = strToU8(file.content);
  }

  const zipped = zipSync(data);
  return zipped;
}

function generateXMLs(worksheets: Worksheet[]) {
  const {
    sharedStringsXml,
    workbookXml,
    workbookXmlRels,
    contentTypesXml,
    stylesXml,
    relsFile,
    themeXml,
    appXml,
    coreXml,
    sheetXmls,
    styleMappers,
  } = createExcelFiles(worksheets);

  const files: { filename: string; content: string }[] = [];
  files.push({ filename: "[Content_Types].xml", content: contentTypesXml });
  files.push({ filename: "_rels/.rels", content: relsFile });
  files.push({ filename: "docProps/app.xml", content: appXml });
  files.push({ filename: "docProps/core.xml", content: coreXml });
  files.push({
    filename: "xl/sharedStrings.xml",
    content: sharedStringsXml ?? "",
  });
  files.push({ filename: "xl/styles.xml", content: stylesXml });
  files.push({ filename: "xl/workbook.xml", content: workbookXml });
  files.push({
    filename: "xl/_rels/workbook.xml.rels",
    content: workbookXmlRels,
  });
  files.push({ filename: "xl/theme/theme1.xml", content: themeXml });

  for (let i = 0; i < sheetXmls.length; i++) {
    files.push({
      filename: `xl/worksheets/sheet${i + 1}.xml`,
      content: sheetXmls[i]!,
    });
  }

  if (styleMappers.worksheetRels.relsLength > 0) {
    const worksheetRelsXml = styleMappers.worksheetRels.makeXML();
    files.push({
      filename: "xl/worksheets/_rels/sheet1.xml.rels",
      content: worksheetRelsXml,
    });
  }

  return files;
}

function createExcelFiles(worksheets: Worksheet[]) {
  if (worksheets.length === 0) {
    throw new Error("worksheets is empty");
  }

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

  const sheetXmls: string[] = [];
  const worksheetsLength = worksheets.length;
  let count = 0;
  for (const worksheet of worksheets) {
    const defaultColWidth = worksheet.props.defaultColWidth;
    const defaultRowHeight = worksheet.props.defaultRowHeight;
    const sheetData = worksheet.sheetData;
    const xlsxCols = combineColProps(worksheet.cols).map((col) =>
      convertCombinedColToXlsxCol(col, styleMappers, defaultColWidth)
    );
    const xlsxRows = combineRowProps(worksheet.rows).map((row) =>
      convRowToXlsxRow(row, styleMappers)
    );
    const colsXml = makeColsXml(xlsxCols, defaultColWidth);
    const mergeCellsXml = makeMergeCellsXml(worksheet.mergeCells);
    const sheetDataXml = makeSheetDataXml(
      sheetData,
      styleMappers,
      xlsxCols,
      xlsxRows
    );
    const dimension = getDimension(sheetData);
    const tabSelected = count === 0;
    const sheetViewsXml = makeSheetViewsXml(
      tabSelected,
      dimension,
      worksheet.freezePane
    );
    const shhetFormatPrXML = makeSheetFormatPrXml(
      defaultRowHeight,
      defaultColWidth
    );

    // Perhaps passing a UUID to every sheet won't cause any issues,
    // but for the sake of integration testing, only the first sheet is given specific UUID.
    const uuid =
      count === 0 ? "00000000-0001-0000-0000-000000000000" : uuidv4();
    const sheetXml = makeSheetXml(
      uuid,
      colsXml,
      sheetViewsXml,
      shhetFormatPrXML,
      sheetDataXml,
      mergeCellsXml,
      dimension,
      styleMappers.hyperlinks
    );
    sheetXmls.push(sheetXml);
    count++;
  }

  const sharedStringsXml = makeSharedStringsXml(styleMappers.sharedStrings);
  const hasSharedStrings = sharedStringsXml !== null;
  const workbookXml = makeWorkbookXml(worksheets);
  const workbookXmlRels = makeWorkbookXmlRels(
    hasSharedStrings,
    worksheetsLength
  );
  const contentTypesXml = makeContentTypesXml(
    hasSharedStrings,
    worksheetsLength
  );

  const stylesXml = makeStylesXml(styleMappers);
  const relsFile = makeRelsFile();
  const themeXml = makeThemeXml();
  const appXml = makeAppXml();
  const coreXml = makeCoreXml();
  return {
    sharedStringsXml,
    workbookXml,
    workbookXmlRels,
    contentTypesXml,
    stylesXml,
    relsFile,
    themeXml,
    appXml,
    coreXml,
    sheetXmls,
    styleMappers,
  };
}

export function convertCombinedColToXlsxCol(
  col: CombinedCol,
  mappers: StyleMappers,
  defaultWidth: number
): XlsxCol {
  let cellXfId: number | null = null;
  if (col.style) {
    const style = composeXlsxCellStyle(col.style, mappers);
    if (style === null) {
      throw new Error("style is null");
    }
    cellXfId = mappers.cellXfs.getCellXfId(style);
  }

  return {
    min: col.startIndex + 1,
    max: col.endIndex + 1,
    width: col.width ?? defaultWidth,
    customWidth: col.width !== undefined && col.width !== defaultWidth,
    cellXfId: cellXfId,
  };
}

export function convRowToXlsxRow(
  row: CombinedRow,
  styleMappers: StyleMappers
): XlsxRow {
  let cellXfId: number | null = null;
  if (row.style) {
    const style = composeXlsxCellStyle(row.style, styleMappers);
    if (style === null) {
      throw new Error("style is null");
    }
    cellXfId = styleMappers.cellXfs.getCellXfId(style);
  }

  return {
    index: row.index,
    height: row.height ?? DEFAULT_ROW_HEIGHT,
    customHeight: row.height !== undefined && row.height !== DEFAULT_ROW_HEIGHT,
    cellXfId: cellXfId,
  };
}

export function makeColsXml(cols: XlsxCol[], defaultColWidth: number): string {
  if (cols.length === 0) {
    return "";
  }

  let result = "<cols>";
  for (const col of cols) {
    result += `<col min="${col.min}" max="${col.max}"`;

    if (col.customWidth) {
      result += ` width="${col.width}" customWidth="1"`;
    } else {
      result += ` width="${defaultColWidth}"`;
    }

    if (col.cellXfId) {
      result += ` style="${col.cellXfId}"`;
    }

    result += "/>";
  }
  result += "</cols>";

  return result;
}

export function makeMergeCellsXml(mergeCells: MergeCell[]) {
  if (mergeCells.length === 0) {
    return "";
  }

  let result = `<mergeCells count="${mergeCells.length}">`;
  for (const mergeCell of mergeCells) {
    result += `<mergeCell ref="${mergeCell.ref}"/>`;
  }
  result += "</mergeCells>";

  return result;
}

// <sheetViews>
// <sheetView tabSelected="1" workbookViewId="0">
//     <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
//     <selection pane="topRight"/>
// </sheetView>
// </sheetViews>
export function makeSheetViewsXml(
  tabSelected: boolean,
  dimension: { start: string; end: string },
  freezePane: FreezePane | null
) {
  const openingTabSelectedTag = tabSelected
    ? `<sheetView tabSelected="1" workbookViewId="0">`
    : `<sheetView workbookViewId="0">`;

  if (freezePane === null) {
    let result =
      "<sheetViews>" +
      openingTabSelectedTag +
      `<selection activeCell="${dimension.start}" sqref="${dimension.start}"/>` +
      "</sheetView>" +
      "</sheetViews>";
    return result;
  }

  switch (freezePane.target) {
    case "column": {
      let result =
        "<sheetViews>" +
        openingTabSelectedTag +
        `<pane ySplit="${freezePane.split}" topLeftCell="A${
          freezePane.split + 1
        }" activePane="bottomLeft" state="frozen"/>` +
        `<selection pane="bottomLeft" activeCell="${dimension.start}" sqref="${dimension.start}"/>` +
        "</sheetView>" +
        "</sheetViews>";
      return result;
    }
    case "row": {
      let result =
        "<sheetViews>" +
        openingTabSelectedTag +
        `<pane xSplit="${
          freezePane.split
        }" topLeftCell="${convColIndexToColName(
          freezePane.split
        )}1" activePane="topRight" state="frozen"/>` +
        `<selection pane="topRight" activeCell="${dimension.start}" sqref="${dimension.start}"/>` +
        "</sheetView>" +
        "</sheetViews>";
      return result;
    }
    default: {
      const _exhaustiveCheck: never = freezePane.target;
      throw new Error(`unknown freezePane type: ${_exhaustiveCheck}`);
    }
  }
}

export function makeSheetFormatPrXml(
  defaultRowHeight: number,
  defaultColWidth: number
) {
  // There should be no issue with always the defaultColWidth,
  // but due to differences in integration tests with files created in Online Excel,
  // we deliberately avoid adding it when it's the same value as DEFAULT_COL_WIDTH.
  const shhetFormatPrXML =
    defaultColWidth === DEFAULT_COL_WIDTH
      ? `<sheetFormatPr defaultRowHeight="${defaultRowHeight}"/>`
      : `<sheetFormatPr defaultRowHeight="${defaultRowHeight}" defaultColWidth="${defaultColWidth}"/>`;

  return shhetFormatPrXML;
}

export function makeSheetXml(
  uuid: string,
  colsXml: string,
  sheetViewsXml: string,
  sheetFormatPrXml: string,
  sheetDataString: string,
  mergeCellsXml: string,
  dimension: { start: string; end: string },
  hyperlinks: Hyperlinks
) {
  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{${uuid}}">` +
    `<dimension ref="${dimension.start}:${dimension.end}"/>` +
    sheetViewsXml +
    sheetFormatPrXml +
    colsXml +
    sheetDataString;

  if (hyperlinks.getHyperlinks().length > 0) {
    result += hyperlinks.makeXML();
  }

  result +=
    mergeCellsXml +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';

  return result;
}

export function makeSharedStringsXml(sharedStrings: SharedStrings) {
  if (sharedStrings.count === 0) {
    return null;
  }

  let result = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  result += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.count}" uniqueCount="${sharedStrings.uniqueCount}">`;
  for (const str of sharedStrings.getValuesInOrder()) {
    result += `<si><t>${str}</t></si>`;
  }
  result += `</sst>`;
  return result;
}

function findFirstNotBlankRow(sheetData: SheetData) {
  let index = 0;
  for (let i = 0; i < sheetData.length; i++) {
    const row = sheetData[i]!;
    if (row.length > 0) {
      index = i;
      break;
    }
  }
  return index;
}

function findLastNotBlankRow(sheetData: SheetData) {
  let index = 0;
  for (let i = sheetData.length - 1; i >= 0; i--) {
    const row = sheetData[i]!;
    if (row.length > 0) {
      index = i;
      break;
    }
  }
  return index;
}

export function getDimension(sheetData: SheetData) {
  // FIXME: The dimension is alse affected by 'cols'. It can have the correct value even without sheetData.
  if (sheetData.length === 0) {
    // This is a workaround for the case where sheetData is empty.
    return { start: "A1", end: "A1" };
  }

  const firstRowIndex = findFirstNotBlankRow(sheetData);
  const lastRowIndex = findLastNotBlankRow(sheetData);
  if (firstRowIndex === null || lastRowIndex === null) {
    throw new Error("sheetData is empty");
  }

  const firstRowNumber = firstRowIndex + 1;
  const lastRowNumber = lastRowIndex + 1;

  const spans = getSpansFromSheetData(sheetData);
  const { startNumber, endNumber } = spans;
  const firstColumn = convColIndexToColName(startNumber - 1);
  const lastColumn = convColIndexToColName(endNumber - 1);

  return {
    start: `${firstColumn}${firstRowNumber}`,
    end: `${lastColumn}${lastRowNumber}`,
  };
}

export function makeSheetDataXml(
  sheetData: SheetData,
  styleMappers: StyleMappers,
  xlsxCols: XlsxCol[],
  xlsxRows: XlsxRow[]
) {
  const { startNumber, endNumber } = getSpansFromSheetData(sheetData);

  let result = `<sheetData>`;
  let rowIndex = 0;
  for (const row of sheetData) {
    const str = rowToString(
      row,
      rowIndex,
      startNumber,
      endNumber,
      styleMappers,
      xlsxCols,
      xlsxRows
    );
    if (str !== null) {
      result += str;
    }
    rowIndex++;
  }
  result += `</sheetData>`;
  return result;
}

/**
 * <row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
 */
export function rowToString(
  row: RowData,
  rowIndex: number,
  startNumber: number,
  endNumber: number,
  styleMappers: StyleMappers,
  xlsxCols: XlsxCol[],
  xlsxRows: XlsxRow[]
): string | null {
  if (row.length === 0) {
    return null;
  }

  const rowNumber = rowIndex + 1;

  let result = `<row r="${rowNumber}" spans="${startNumber}:${endNumber}"`;

  const xlsxRow = xlsxRows.find((it) => it.index === rowIndex);
  if (xlsxRow) {
    if (xlsxRow.cellXfId) {
      result += ` s="${xlsxRow.cellXfId}" customFormat="1"`;
    }

    if (xlsxRow.height && xlsxRow.customHeight) {
      result += ` ht="${xlsxRow.height}" customHeight="1"`;
    }
  }
  result += ">";

  let columnIndex = 0;
  for (const cell of row) {
    if (cell !== null) {
      result += makeCellXml(
        convertCellToXlsxCell(
          cell,
          columnIndex,
          rowIndex,
          styleMappers,
          xlsxCols,
          xlsxRows
        )
      );
    }

    columnIndex++;
  }

  result += `</row>`;
  return result;
}

export function getSpansFromSheetData(sheetData: SheetData) {
  const all = sheetData
    .map((row) => {
      return getSpans(row);
    })
    .filter((row) => row !== null) as {
    startNumber: number;
    endNumber: number;
  }[];
  const minStartNumber = Math.min(...all.map((row) => row.startNumber));
  const maxEndNumber = Math.max(...all.map((row) => row.endNumber));
  return { startNumber: minStartNumber, endNumber: maxEndNumber };
}

export function findFirstNonNullCell(row: RowData) {
  let index = 0;
  let firstNonNullCell = null;
  for (let i = 0; i < row.length; i++) {
    if (row[i] !== null) {
      firstNonNullCell = row[i]!;
      index = i;
      break;
    }
  }
  return { firstNonNullCell, index };
}

/**
 *  [null, null, null, nonnull, null] => index is 3
 */
export function findLastNonNullCell(row: RowData) {
  let index = 0;
  let lastNonNullCell = null;
  for (let i = row.length - 1; i >= 0; i--) {
    if (row[i] !== null) {
      lastNonNullCell = row[i]!;
      index = i;
      break;
    }
  }
  return { lastNonNullCell, index };
}

export function getSpans(row: RowData) {
  const first = findFirstNonNullCell(row);
  if (first === undefined || first === null) {
    return null;
  }

  const last = findLastNonNullCell(row);
  if (last === undefined) {
    return null;
  }

  const startNumber = first.index + 1;
  const endNumber = last.index + 1;

  return { startNumber, endNumber };
}

/**
 * https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
 * @param isoString
 * @returns
 */
export function convertIsoStringToSerialValue(isoString: string): number {
  const baseDate = new Date("1899-12-31T00:00:00.000Z");
  const targetDate = new Date(isoString);
  const differenceInDays =
    (targetDate.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24);
  // Excel uses January 0, 1900 as a base (which is actually December 31, 1899), so add 1 to the result
  return differenceInDays + 1;
}

function assignDateStyleIfUndefined(cell: Cell) {
  if (cell.type === "date" && cell.style === undefined) {
    cell.style = { numberFormat: { formatCode: "yyyy-mm-dd" } };
  }
}

function assignHyperlinkStyleIfUndefined(cell: Cell) {
  if (cell.type === "hyperlink" && cell.style === undefined) {
    cell.style = {
      font: {
        name: "Calibri",
        size: 11,
        color: "0563c1",
        underline: "single",
      },
    };
  }
}

export function composeXlsxCellStyle(
  style: CellStyle | undefined,
  mappers: StyleMappers
): XlsxCellStyle | null {
  if (style) {
    const _style: XlsxCellStyle = {
      fillId: style.fill ? mappers.fills.getFillId(style.fill) : 0,
      fontId: style.font ? mappers.fonts.getFontId(style.font) : 0,
      borderId: style.border ? mappers.borders.getBorderId(style.border) : 0,
      numFmtId: style.numberFormat
        ? mappers.numberFormats.getNumFmtId(style.numberFormat.formatCode)
        : 0,
    };

    if (style.alignment) {
      _style.alignment = style.alignment;
    }

    return _style;
  }

  return null;
}

function getCellXfId(
  cell: Cell,
  colName: string,
  rowIndex: number,
  styleMappers: StyleMappers,
  xlsxCols: XlsxCol[],
  xlsxRows: XlsxRow[]
) {
  const composedStyle = composeXlsxCellStyle(cell.style, styleMappers);
  if (composedStyle) {
    return styleMappers.cellXfs.getCellXfId(composedStyle);
  }

  const foundCol = xlsxCols.find((it) => isInRange(colName, it.min, it.max));
  if (foundCol) {
    return foundCol.cellXfId;
  }

  const foundRow = xlsxRows.find((it) => it.index === rowIndex);
  if (foundRow) {
    return foundRow.cellXfId;
  }

  return null;
}

export function convertCellToXlsxCell(
  cell: Cell,
  columnIndex: number,
  rowIndex: number,
  styleMappers: StyleMappers,
  xlsxCols: XlsxCol[],
  xlsxRows: XlsxRow[]
): XlsxCell {
  const rowNumber = rowIndex + 1;
  const colName = convColIndexToColName(columnIndex);

  switch (cell.type) {
    case "number": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        rowIndex,
        styleMappers,
        xlsxCols,
        xlsxRows
      );
      return {
        type: "number",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "string": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        rowIndex,
        styleMappers,
        xlsxCols,
        xlsxRows
      );
      const sharedStringId = styleMappers.sharedStrings.getIndex(cell.value);
      return {
        type: "string",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        sharedStringId: sharedStringId,
        cellXfId: cellXfId,
      };
    }
    case "date": {
      assignDateStyleIfUndefined(cell);
      const cellXfId = getCellXfId(
        cell,
        colName,
        rowIndex,
        styleMappers,
        xlsxCols,
        xlsxRows
      );
      return {
        type: "date",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "hyperlink": {
      assignHyperlinkStyleIfUndefined(cell);
      const composedStyle = composeXlsxCellStyle(cell.style, styleMappers);
      if (composedStyle === null) {
        throw new Error("composedStyle is null for hyperlink");
      }
      const xfId = styleMappers.cellStyleXfs.getCellStyleXfId(composedStyle);
      if (xfId === null) {
        throw new Error("xfId is null for hyperlink");
      }

      const cellXf: CellXf = {
        xfId: xfId,
        ...composedStyle,
      };
      const cellXfId = styleMappers.cellXfs.getCellXfId(cellXf);
      const sharedStringId = styleMappers.sharedStrings.getIndex(cell.text);

      styleMappers.cellStyles.getCellStyleId({
        name: "Hyperlink",
        xfId: xfId,
        uid: "{00000000-000B-0000-0000-000008000000}",
      });

      if (cell.linkType === "external") {
        const rid = styleMappers.worksheetRels.addWorksheetRel(cell.value);
        styleMappers.hyperlinks.addHyperlink({
          linkType: "external",
          ref: `${colName}${rowNumber}`,
          rid: rid,
          uuid: uuidv4(),
        });
      } else if (cell.linkType === "internal") {
        styleMappers.hyperlinks.addHyperlink({
          linkType: "internal",
          ref: `${colName}${rowNumber}`,
          location: cell.value,
          display: cell.text,
          uuid: uuidv4(),
        });
      } else if (cell.linkType === "email") {
        const rid = styleMappers.worksheetRels.addWorksheetRel(
          `mailto:${cell.value}`
        );
        styleMappers.hyperlinks.addHyperlink({
          linkType: "email",
          ref: `${colName}${rowNumber}`,
          rid: rid,
          uuid: uuidv4(),
        });
      }

      return {
        type: "hyperlink",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        sharedStringId: sharedStringId,
        cellXfId: cellXfId,
      };
    }
    case "boolean": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        rowIndex,
        styleMappers,
        xlsxCols,
        xlsxRows
      );
      return {
        type: "boolean",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "formula": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        rowIndex,
        styleMappers,
        xlsxCols,
        xlsxRows
      );
      return {
        type: "formula",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "merged": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        rowIndex,
        styleMappers,
        xlsxCols,
        xlsxRows
      );
      return {
        type: "merged",
        colName: colName,
        rowNumber: rowNumber,
        cellXfId: cellXfId,
      };
    }
    default: {
      const _exhaustiveCheck: never = cell;
      throw new Error(`unknown cell type: ${_exhaustiveCheck}`);
    }
  }
}

export function makeCellXml(cell: XlsxCell) {
  switch (cell.type) {
    case "number": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s}><v>${cell.value}</v></c>`;
    }
    case "string": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s} t="s"><v>${cell.sharedStringId}</v></c>`;
    }
    case "date": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      const serialValue = convertIsoStringToSerialValue(cell.value);
      return `<c r="${cell.colName}${cell.rowNumber}"${s}><v>${serialValue}</v></c>`;
    }
    case "hyperlink": {
      const s = ` s="${cell.cellXfId}"`;
      return `<c r="${cell.colName}${cell.rowNumber}"${s} t="s"><v>${cell.sharedStringId}</v></c>`;
    }
    case "boolean": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      const v = cell.value ? 1 : 0;
      return `<c r="${cell.colName}${cell.rowNumber}"${s} t="b"><v>${v}</v></c>`;
    }
    case "formula": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s}><f>${cell.value}</f></c>`;
    }
    case "merged": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s}/>`;
    }
    default: {
      const _exhaustiveCheck: never = cell;
      throw new Error(`unknown cell type: ${_exhaustiveCheck}`);
    }
  }
}

function makeWorkbookXmlRels(
  sharedStrings: boolean,
  wooksheetsLength: number
): string {
  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

  let index = 1;
  while (index <= wooksheetsLength) {
    result += `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index}.xml"/>`;
    index++;
  }

  result += `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`;
  index++;

  result += `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  index++;

  if (sharedStrings) {
    result += `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
  }

  result += "</Relationships>";
  return result;
}

function makeCoreXml() {
  const isoDate = new Date().toISOString();

  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:title></dc:title>" +
    "<dc:subject></dc:subject>" +
    "<dc:creator></dc:creator>" +
    "<cp:keywords></cp:keywords>" +
    "<dc:description></dc:description>" +
    "<cp:lastModifiedBy></cp:lastModifiedBy>" +
    "<cp:revision></cp:revision>" +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:modified><cp:category></cp:category>` +
    "<cp:contentStatus></cp:contentStatus>" +
    "</cp:coreProperties>";

  return result;
}

function makeStylesXml(styleMappers: StyleMappers) {
  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
    styleMappers.numberFormats.makeXml() +
    // results.push('<fonts count="1">');
    // results.push("<font>");
    // results.push('<sz val="11"/>');
    // results.push('<color theme="1"/>');
    // results.push('<name val="Calibri"/>');
    // results.push('<family val="2"/>');
    // results.push('<scheme val="minor"/></font>');
    // results.push("</fonts>");
    styleMappers.fonts.makeXml() +
    // results.push('<fills count="2">');
    // results.push('<fill><patternFill patternType="none"/></fill>');
    // results.push('<fill><patternFill patternType="gray125"/></fill>');
    // results.push("</fills>");
    styleMappers.fills.makeXml() +
    // results.push('<borders count="1">');
    // results.push("<border><left/><right/><top/><bottom/><diagonal/></border>");
    // results.push("</borders>");
    styleMappers.borders.makeXml() +
    // results.push(
    //   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    // );
    styleMappers.cellStyleXfs.makeXml() +
    // results.push(
    //   '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    // );
    styleMappers.cellXfs.makeXml() +
    // results.push(
    //   '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>'
    // );
    styleMappers.cellStyles.makeXml() +
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>' +
    "<extLst>" +
    '<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>' +
    '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
    '<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext>' +
    "</extLst>" +
    "</styleSheet>";

  return result;
}

function makeWorkbookXml(worksheets: Worksheet[]) {
  const documentId = uuidv4();

  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">' +
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="27123"/>' +
    '<workbookPr defaultThemeVersion="166925"/>' +
    `<xr:revisionPtr revIDLastSave="0" documentId="8_{${documentId}}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>` +
    "<bookViews>" +
    '<workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/>' +
    "</bookViews>" +
    "<sheets>";

  let sheetId = 1;
  for (const sheet of worksheets) {
    result += `<sheet name="${sheet.name}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`;
    sheetId++;
  }

  result +=
    "</sheets>" +
    "<extLst>" +
    '<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
    '<x15:workbookPr chartTrackingRefBase="1"/>' +
    "</ext>" +
    '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">' +
    "<xcalcf:calcFeatures>" +
    '<xcalcf:feature name="microsoft.com:RD"/>' +
    '<xcalcf:feature name="microsoft.com:Single"/>' +
    '<xcalcf:feature name="microsoft.com:FV"/>' +
    '<xcalcf:feature name="microsoft.com:CNMTM"/>' +
    '<xcalcf:feature name="microsoft.com:LET_WF"/>' +
    '<xcalcf:feature name="microsoft.com:LAMBDA_WF"/>' +
    '<xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/>' +
    "</xcalcf:calcFeatures>" +
    "</ext>" +
    "</extLst>" +
    "</workbook>";

  return result;
}

function makeAppXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    "<Application>xlsx-kaku</Application>" +
    "<Manager></Manager>" +
    "<Company></Company>" +
    "<HyperlinkBase></HyperlinkBase>" +
    "<AppVersion>16.0300</AppVersion>" +
    "</Properties>"
  );
}

function makeRelsFile() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    "</Relationships>"
  );
}

function makeContentTypesXml(sharedStrings: boolean, sheetsLength: number) {
  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';

  for (let i = 1; i <= sheetsLength; i++) {
    result += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }

  result +=
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';

  if (sharedStrings) {
    result +=
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
  }

  result += "</Types>";

  return result;
}
