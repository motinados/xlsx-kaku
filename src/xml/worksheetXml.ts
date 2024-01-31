import { v4 as uuidv4 } from "uuid";
import { FreezePane, MergeCell, Worksheet } from "..";
import { Cell, CellStyle, RowData, SheetData } from "../sheetData";
import { StyleMappers } from "../writer";
import { convColIndexToColName, convColNameToColIndex } from "../utils";
import { Alignment, CellXf } from "../cellXfs";
import { Hyperlinks } from "../hyperlinks";
import {
  ColProps,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  RowProps,
} from "../worksheet";

export type XlsxCol = {
  /** e.g. column A is 0 */
  index: number;
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

export type GroupedXlsxCol = {
  startIndex: number;
  endIndex: number;
  width: number;
  customWidth: boolean;
  cellXfId: number | null;
};

export function makeWorksheetXml(
  worksheet: Worksheet,
  styleMappers: StyleMappers,
  sheetCnt: number
) {
  styleMappers.hyperlinks.reset();
  styleMappers.worksheetRels.reset();

  const defaultColWidth = worksheet.props.defaultColWidth;
  const defaultRowHeight = worksheet.props.defaultRowHeight;
  const sheetData = worksheet.sheetData;

  const xlsxCols = new Map<number, XlsxCol>();
  for (const col of worksheet.cols.values()) {
    const xlsxCol = createXlsxColFromColProps(
      col,
      styleMappers,
      defaultColWidth
    );
    xlsxCols.set(xlsxCol.index, xlsxCol);
  }

  const xlsxRows = new Map<number, XlsxRow>();
  for (const row of worksheet.rows.values()) {
    const xlsxRow = createXlsxRowFromRowProps(row, styleMappers);
    xlsxRows.set(xlsxRow.index, xlsxRow);
  }

  const colsXml = makeColsXml(groupXlsxCols(xlsxCols), defaultColWidth);
  const mergeCellsXml = makeMergeCellsXml(worksheet.mergeCells);
  const sheetDataXml = makeSheetDataXml(
    sheetData,
    styleMappers,
    xlsxCols,
    xlsxRows
  );
  const dimension = getDimension(sheetData);
  const tabSelected = sheetCnt === 0;
  const sheetViewsXml = makeSheetViewsXml(
    tabSelected,
    dimension,
    worksheet.freezePane
  );
  const sheetFormatPrXML = makeSheetFormatPrXml(
    defaultRowHeight,
    defaultColWidth
  );

  // Perhaps passing a UUID to every sheet won't cause any issues,
  // but for the sake of integration testing, only the first sheet is given specific UUID.
  const uuid =
    sheetCnt === 0 ? "00000000-0001-0000-0000-000000000000" : uuidv4();
  const sheetXml = composeSheetXml(
    uuid,
    colsXml,
    sheetViewsXml,
    sheetFormatPrXML,
    sheetDataXml,
    mergeCellsXml,
    dimension,
    styleMappers.hyperlinks
  );

  let worksheetRels;
  if (styleMappers.worksheetRels.relsLength > 0) {
    worksheetRels = styleMappers.worksheetRels.makeXML();
  } else {
    worksheetRels = null;
  }

  return {
    sheetXml,
    worksheetRels,
  };
}

export function createXlsxColFromColProps(
  col: ColProps,
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
    index: col.index,
    width: col.width ?? defaultWidth,
    customWidth: col.width !== undefined && col.width !== defaultWidth,
    cellXfId: cellXfId,
  };
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

export function createXlsxRowFromRowProps(
  row: RowProps,
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

export function isEqualsXlsxCol(a: XlsxCol, b: XlsxCol) {
  return (
    // index must not be compared.
    a.width === b.width &&
    a.customWidth === b.customWidth &&
    a.cellXfId === b.cellXfId
  );
}

export function groupXlsxCols(cols: Map<number, XlsxCol>) {
  const result: GroupedXlsxCol[] = [];
  let startCol: XlsxCol;
  let endCol: XlsxCol;

  let i = 0;
  for (const col of cols.values()) {
    if (i === 0) {
      // the first
      startCol = col;
      endCol = col;
    } else {
      if (isEqualsXlsxCol(startCol!, col)) {
        endCol = col;
      } else {
        result.push({
          startIndex: startCol!.index,
          endIndex: endCol!.index,
          width: startCol!.width,
          customWidth: startCol!.customWidth,
          cellXfId: startCol!.cellXfId,
        });
        startCol = col;
        endCol = col;
      }
    }

    if (i == cols.size - 1) {
      // the last
      result.push({
        startIndex: startCol!.index,
        endIndex: endCol!.index,
        width: startCol!.width,
        customWidth: startCol!.customWidth,
        cellXfId: startCol!.cellXfId,
      });
    }

    i++;
  }

  return result;
}

export function makeColsXml(
  cols: GroupedXlsxCol[],
  defaultColWidth: number
): string {
  if (cols.length === 0) {
    return "";
  }

  let result = "<cols>";
  for (const col of cols) {
    result += `<col min="${col.startIndex + 1}" max="${col.endIndex + 1}"`;

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

export function makeSheetDataXml(
  sheetData: SheetData,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  xlsxRows: Map<number, XlsxRow>
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

/**
 * <row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
 */
export function rowToString(
  row: RowData,
  rowIndex: number,
  startNumber: number,
  endNumber: number,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  xlsxRows: Map<number, XlsxRow>
): string | null {
  if (row.length === 0) {
    return null;
  }

  const rowNumber = rowIndex + 1;

  let result = `<row r="${rowNumber}" spans="${startNumber}:${endNumber}"`;

  const xlsxRow = xlsxRows.get(rowIndex);
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
          xlsxRow
        )
      );
    }

    columnIndex++;
  }

  result += `</row>`;
  return result;
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

export function convertCellToXlsxCell(
  cell: Cell,
  columnIndex: number,
  rowIndex: number,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  xlsxRow: XlsxRow | undefined
): XlsxCell {
  const rowNumber = rowIndex + 1;
  const colName = convColIndexToColName(columnIndex);

  switch (cell.type) {
    case "number": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
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
        styleMappers,
        xlsxCols,
        xlsxRow
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
        styleMappers,
        xlsxCols,
        xlsxRow
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
        styleMappers,
        xlsxCols,
        xlsxRow
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
        styleMappers,
        xlsxCols,
        xlsxRow
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
        styleMappers,
        xlsxCols,
        xlsxRow
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

function getCellXfId(
  cell: Cell,
  colName: string,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  foundRow: XlsxRow | undefined
) {
  const composedStyle = composeXlsxCellStyle(cell.style, styleMappers);
  if (composedStyle) {
    return styleMappers.cellXfs.getCellXfId(composedStyle);
  }

  const foundCol = xlsxCols.get(convColNameToColIndex(colName));
  if (foundCol) {
    return foundCol.cellXfId;
  }

  if (foundRow) {
    return foundRow.cellXfId;
  }

  return null;
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

export function composeSheetXml(
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
