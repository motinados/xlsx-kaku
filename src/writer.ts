import * as fs from "node:fs";
import path from "node:path";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import { Cell, RowData, SheetData } from "./sheetData";
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
import { Col, FreezePane, MergeCell, Row, Worksheet } from "./worksheet";
import { convNumberToColumn } from "./utils";

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

type XlsxCellStyle = {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: Alignment;
};

export async function writeXlsx(filepath: string, worksheets: Worksheet[]) {
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

  const xlsxPath = path.resolve(filepath);
  const basePath = path.dirname(filepath);
  const workDir = path.join(basePath, "work");
  if (!fs.existsSync(workDir)) {
    fs.mkdirSync(workDir, { recursive: true });
  }
  fs.writeFileSync(path.join(workDir, "[Content_Types].xml"), contentTypesXml);

  const _relsPath = path.resolve(workDir, "_rels");
  if (!fs.existsSync(_relsPath)) {
    fs.mkdirSync(_relsPath, { recursive: true });
  }
  fs.writeFileSync(path.join(_relsPath, ".rels"), relsFile);

  const docPropsPath = path.resolve(workDir, "docProps");
  if (!fs.existsSync(docPropsPath)) {
    fs.mkdirSync(docPropsPath, { recursive: true });
  }
  fs.writeFileSync(path.join(docPropsPath, "app.xml"), appXml);
  fs.writeFileSync(path.join(docPropsPath, "core.xml"), coreXml);

  const xlPath = path.resolve(workDir, "xl");
  if (!fs.existsSync(xlPath)) {
    fs.mkdirSync(xlPath, { recursive: true });
  }
  if (sharedStringsXml !== null) {
    fs.writeFileSync(path.join(xlPath, "sharedStrings.xml"), sharedStringsXml);
  }
  fs.writeFileSync(path.join(xlPath, "styles.xml"), stylesXml);
  fs.writeFileSync(path.join(xlPath, "workbook.xml"), workbookXml);

  const xl_relsPath = path.resolve(xlPath, "_rels");
  if (!fs.existsSync(xl_relsPath)) {
    fs.mkdirSync(xl_relsPath, { recursive: true });
  }
  fs.writeFileSync(
    path.join(xl_relsPath, "workbook.xml.rels"),
    workbookXmlRels
  );

  const themePath = path.resolve(xlPath, "theme");
  if (!fs.existsSync(themePath)) {
    fs.mkdirSync(themePath, { recursive: true });
  }
  fs.writeFileSync(path.join(themePath, "theme1.xml"), themeXml);

  const worksheetsPath = path.resolve(xlPath, "worksheets");
  if (!fs.existsSync(worksheetsPath)) {
    fs.mkdirSync(worksheetsPath, { recursive: true });
  }

  let sheetIndex = 1;
  for (const sheetXml of sheetXmls) {
    fs.writeFileSync(
      path.join(worksheetsPath, `sheet${sheetIndex}.xml`),
      sheetXml
    );
    sheetIndex++;
  }

  if (styleMappers.worksheetRels.relsLength > 0) {
    const worksheets_relsPath = path.resolve(worksheetsPath, "_rels");
    if (!fs.existsSync(worksheets_relsPath)) {
      fs.mkdirSync(worksheets_relsPath, { recursive: true });
    }
    const worksheetRelsXml = styleMappers.worksheetRels.makeXML();
    fs.writeFileSync(
      path.join(worksheets_relsPath, "sheet1.xml.rels"),
      worksheetRelsXml
    );
  }

  await zipToXlsx(workDir, xlsxPath);
  fs.rmSync(workDir, { recursive: true });
}

export function createExcelFiles(worksheets: Worksheet[]) {
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
  for (const worksheet of worksheets) {
    const sheetData = worksheet.sheetData;
    const colsXml = makeColsXml(worksheet.cols);
    const mergeCellsXml = makeMergeCellsXml(worksheet.mergeCells);
    const sheetDataXml = makeSheetDataXml(
      sheetData,
      worksheet.rows,
      styleMappers
    );
    const dimension = getDimension(sheetData);
    const sheetViewsXml = makeSheetViewsXml(dimension, worksheet.freezePane);
    const sheetXml = makeSheetXml(
      colsXml,
      sheetViewsXml,
      sheetDataXml,
      mergeCellsXml,
      dimension,
      styleMappers.hyperlinks
    );
    sheetXmls.push(sheetXml);
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

export function zipToXlsx(sourceDir: string, outPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => {
      resolve();
    });

    archive.on("warning", (err) => {
      if (err.code === "ENOENT") {
        console.warn(err);
      } else {
        reject(err);
      }
    });

    archive.on("error", (err) => {
      reject(err);
    });

    archive.pipe(output);
    archive.directory(sourceDir, false);
    archive.finalize();
  });
}

export function makeColsXml(cols: Col[]): string {
  if (cols.length === 0) {
    return "";
  }

  const results: string[] = [];
  results.push("<cols>");
  for (const col of cols) {
    results.push(
      `<col min="${col.min}" max="${col.max}" width="${col.width}" customWidth="1"/>`
    );
  }
  results.push("</cols>");
  return results.join("");
}

export function makeMergeCellsXml(mergeCells: MergeCell[]) {
  if (mergeCells.length === 0) {
    return "";
  }

  const results: string[] = [];
  results.push(`<mergeCells count="${mergeCells.length}">`);
  for (const mergeCell of mergeCells) {
    results.push(`<mergeCell ref="${mergeCell.ref}"/>`);
  }
  results.push("</mergeCells>");
  return results.join("");
}

// <sheetViews>
// <sheetView tabSelected="1" workbookViewId="0">
//     <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
//     <selection pane="topRight"/>
// </sheetView>
// </sheetViews>
export function makeSheetViewsXml(
  dimension: { start: string; end: string },
  freezePane: FreezePane | null
) {
  if (freezePane === null) {
    let result = "";
    result += "<sheetViews>";
    result += `<sheetView tabSelected="1" workbookViewId="0">`;
    result += `<selection activeCell="${dimension.start}" sqref="${dimension.start}"/>`;
    result += "</sheetView>";
    result += "</sheetViews>";
    return result;
  }

  switch (freezePane.type) {
    case "column": {
      let result = "";
      result += "<sheetViews>";
      result += `<sheetView tabSelected="1" workbookViewId="0">`;
      result += `<pane ySplit="${freezePane.split}" topLeftCell="A${
        freezePane.split + 1
      }" activePane="bottomLeft" state="frozen"/>`;
      result += `<selection pane="bottomLeft" activeCell="${dimension.start}" sqref="${dimension.start}"/>`;
      result += "</sheetView>";
      result += "</sheetViews>";
      return result;
    }
    case "row": {
      let result = "";
      result += "<sheetViews>";
      result += `<sheetView tabSelected="1" workbookViewId="0">`;
      result += `<pane xSplit="${
        freezePane.split
      }" topLeftCell="${convNumberToColumn(
        freezePane.split
      )}1" activePane="topRight" state="frozen"/>`;
      result += `<selection pane="topRight" activeCell="${dimension.start}" sqref="${dimension.start}"/>`;
      result += "</sheetView>";
      result += "</sheetViews>";
      return result;
    }
    default: {
      const _exhaustiveCheck: never = freezePane.type;
      throw new Error(`unknown freezePane type: ${_exhaustiveCheck}`);
    }
  }
}

export function makeSheetXml(
  colsXml: string,
  sheetViewsXml: string,
  sheetDataString: string,
  mergeCellsXml: string,
  dimension: { start: string; end: string },
  hyperlinks: Hyperlinks
) {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">'
  );
  results.push(`<dimension ref="${dimension.start}:${dimension.end}"/>`);
  results.push(sheetViewsXml);
  results.push('<sheetFormatPr defaultRowHeight="13.5"/>');
  results.push(colsXml);
  results.push(sheetDataString);

  if (hyperlinks.getHyperlinks().length > 0) {
    results.push(hyperlinks.makeXML());
  }

  results.push(mergeCellsXml);
  results.push(
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
  );
  return results.join("");
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
  const firstRowIndex = findFirstNotBlankRow(sheetData);
  const lastRowIndex = findLastNotBlankRow(sheetData);
  if (firstRowIndex === null || lastRowIndex === null) {
    throw new Error("sheetData is empty");
  }

  const firstRowNumber = firstRowIndex + 1;
  const lastRowNumber = lastRowIndex + 1;

  const spans = getSpansFromSheetData(sheetData);
  const { startNumber, endNumber } = spans;
  const firstColumn = convNumberToColumn(startNumber - 1);
  const lastColumn = convNumberToColumn(endNumber - 1);

  return {
    start: `${firstColumn}${firstRowNumber}`,
    end: `${lastColumn}${lastRowNumber}`,
  };
}

export function makeSheetDataXml(
  sheetData: SheetData,
  rows: Row[],
  styleMappers: StyleMappers
) {
  const { startNumber, endNumber } = getSpansFromSheetData(sheetData);

  let result = `<sheetData>`;
  let rowIndex = 0;
  for (const row of sheetData) {
    const rowHeight = rows.find((it) => it.index === rowIndex)?.height ?? null;
    const str = rowToString(
      row,
      rowIndex,
      rowHeight,
      startNumber,
      endNumber,
      styleMappers
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
  rowHeight: number | null,
  startNumber: number,
  endNumber: number,
  styleMappers: StyleMappers
): string | null {
  if (row.length === 0) {
    return null;
  }

  const rowNumber = rowIndex + 1;

  let result = rowHeight
    ? `<row r="${rowNumber}" spans="${startNumber}:${endNumber}" ht="${rowHeight}" customHeight="1">`
    : `<row r="${rowNumber}" spans="${startNumber}:${endNumber}">`;

  let columnIndex = 0;
  for (const cell of row) {
    if (cell !== null) {
      result += cellToString(cell, columnIndex, rowIndex, styleMappers);
    }

    columnIndex++;
  }

  result += `</row>`;
  return result;
}

export function getSpansFromSheetData(sheetData: SheetData) {
  const all = sheetData
    .map((row) => {
      const spans = getSpans(row);
      if (spans === null) {
        return null;
      }
      return spans;
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
        underline: true,
      },
    };
  }
}

export function composeXlsxCellStyle(
  cell: Cell,
  mappers: StyleMappers
): XlsxCellStyle | null {
  if (cell.style) {
    const style: XlsxCellStyle = {
      fillId: cell.style.fill ? mappers.fills.getFillId(cell.style.fill) : 0,
      fontId: cell.style.font ? mappers.fonts.getFontId(cell.style.font) : 0,
      borderId: cell.style.border
        ? mappers.borders.getBorderId(cell.style.border)
        : 0,
      numFmtId: cell.style.numberFormat
        ? mappers.numberFormats.getNumFmtId(cell.style.numberFormat.formatCode)
        : 0,
    };

    if (cell.style.alignment) {
      style.alignment = cell.style.alignment;
    }

    return style;
  }

  return null;
}

export function cellToString(
  cell: Cell,
  columnIndex: number,
  rowIndex: number,
  styleMappers: StyleMappers
) {
  const rowNumber = rowIndex + 1;
  const column = convNumberToColumn(columnIndex);

  assignDateStyleIfUndefined(cell);
  assignHyperlinkStyleIfUndefined(cell);
  const composedStyle = composeXlsxCellStyle(cell, styleMappers);

  switch (cell.type) {
    case "number": {
      const cellXfId = composedStyle
        ? styleMappers.cellXfs.getCellXfId(composedStyle)
        : null;
      const s = cellXfId ? ` s="${cellXfId}"` : "";
      return `<c r="${column}${rowNumber}"${s}><v>${cell.value}</v></c>`;
    }
    case "string": {
      const cellXfId = composedStyle
        ? styleMappers.cellXfs.getCellXfId(composedStyle)
        : null;
      const s = cellXfId ? ` s="${cellXfId}"` : "";
      const index = styleMappers.sharedStrings.getIndex(cell.value);
      return `<c r="${column}${rowNumber}"${s} t="s"><v>${index}</v></c>`;
    }
    case "date": {
      const cellXfId = composedStyle
        ? styleMappers.cellXfs.getCellXfId(composedStyle)
        : null;
      const s = cellXfId ? ` s="${cellXfId}"` : "";
      const serialValue = convertIsoStringToSerialValue(cell.value);
      return `<c r="${column}${rowNumber}"${s}><v>${serialValue}</v></c>`;
    }
    case "hyperlink": {
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
      const s = ` s="${cellXfId}"`;
      const index = styleMappers.sharedStrings.getIndex(cell.value);

      styleMappers.cellStyles.getCellStyleId({
        name: "Hyperlink",
        xfId: xfId,
        uid: "{00000000-000B-0000-0000-000008000000}",
      });

      const rid = styleMappers.worksheetRels.addWorksheetRel(cell.value);

      styleMappers.hyperlinks.addHyperlink({
        ref: `${column}${rowNumber}`,
        rid: rid,
        uuid: uuidv4(),
      });

      return `<c r="${column}${rowNumber}"${s} t="s"><v>${index}</v></c>`;
    }
    case "boolean": {
      const cellXfId = composedStyle
        ? styleMappers.cellXfs.getCellXfId(composedStyle)
        : null;
      const s = cellXfId ? ` s="${cellXfId}"` : "";
      const v = cell.value ? 1 : 0;
      return `<c r="${column}${rowNumber}"${s} t="b"><v>${v}</v></c>`;
    }
    case "merged": {
      const cellXfId = composedStyle
        ? styleMappers.cellXfs.getCellXfId(composedStyle)
        : null;
      const s = cellXfId ? ` s="${cellXfId}"` : "";
      return `<c r="${column}${rowNumber}"${s}/>`;
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
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  );

  let index = 1;
  while (index <= wooksheetsLength) {
    results.push(
      `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index}.xml"/>`
    );
    index++;
  }

  results.push(
    `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`
  );
  index++;

  results.push(
    `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  );
  index++;

  if (sharedStrings) {
    results.push(
      `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`
    );
  }
  results.push("</Relationships>");
  return results.join("");
}

function makeCoreXml() {
  const isoDate = new Date().toISOString();
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
  );
  results.push("<dc:title></dc:title>");
  results.push("<dc:subject></dc:subject>");
  results.push("<dc:creator></dc:creator>");
  results.push("<cp:keywords></cp:keywords>");
  results.push("<dc:description></dc:description>");
  results.push("<cp:lastModifiedBy></cp:lastModifiedBy>");
  results.push("<cp:revision></cp:revision>");
  results.push(
    `<dcterms:created xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:created>`
  );
  results.push(
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:modified><cp:category></cp:category>`
  );
  results.push("<cp:contentStatus></cp:contentStatus>");
  results.push("</cp:coreProperties>");
  return results.join("");
}

function makeStylesXml(styleMappers: StyleMappers) {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">'
  );

  results.push(styleMappers.numberFormats.makeXml());

  // results.push('<fonts count="1">');
  // results.push("<font>");
  // results.push('<sz val="11"/>');
  // results.push('<color theme="1"/>');
  // results.push('<name val="Calibri"/>');
  // results.push('<family val="2"/>');
  // results.push('<scheme val="minor"/></font>');
  // results.push("</fonts>");
  results.push(styleMappers.fonts.makeXml());

  // results.push('<fills count="2">');
  // results.push('<fill><patternFill patternType="none"/></fill>');
  // results.push('<fill><patternFill patternType="gray125"/></fill>');
  // results.push("</fills>");
  results.push(styleMappers.fills.makeXml());

  // results.push('<borders count="1">');
  // results.push("<border><left/><right/><top/><bottom/><diagonal/></border>");
  // results.push("</borders>");
  results.push(styleMappers.borders.makeXml());

  // results.push(
  //   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  // );
  results.push(styleMappers.cellStyleXfs.makeXml());

  // results.push(
  //   '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  // );
  results.push(styleMappers.cellXfs.makeXml());

  // results.push(
  //   '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>'
  // );
  results.push(styleMappers.cellStyles.makeXml());

  results.push(
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>'
  );
  results.push("<extLst>");
  results.push(
    '<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
  );
  results.push(
    '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>'
  );
  results.push(
    '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
  );
  results.push(
    '<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext>'
  );
  results.push("</extLst>");
  results.push("</styleSheet>");
  return results.join("");
}

function makeWorkbookXml(worksheets: Worksheet[]) {
  const documentId = uuidv4();
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">'
  );
  results.push(
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="27123"/>'
  );
  results.push('<workbookPr defaultThemeVersion="166925"/>');
  results.push(
    `<xr:revisionPtr revIDLastSave="0" documentId="8_{${documentId}}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>`
  );
  results.push("<bookViews>");
  results.push(
    '<workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/>'
  );
  results.push("</bookViews>");

  results.push("<sheets>");
  let sheetId = 1;
  for (const sheet of worksheets) {
    // results.push('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>');
    results.push(
      `<sheet name="${sheet.name}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`
    );
    sheetId++;
  }
  results.push("</sheets>");

  results.push('<calcPr calcId="191028"/>');
  results.push("<extLst>");
  results.push(
    '<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
  );
  results.push('<x15:workbookPr chartTrackingRefBase="1"/>');
  results.push("</ext>");
  results.push(
    '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">'
  );
  results.push("<xcalcf:calcFeatures>");
  results.push('<xcalcf:feature name="microsoft.com:RD"/>');
  results.push('<xcalcf:feature name="microsoft.com:Single"/>');
  results.push('<xcalcf:feature name="microsoft.com:FV"/>');
  results.push('<xcalcf:feature name="microsoft.com:CNMTM"/>');
  results.push('<xcalcf:feature name="microsoft.com:LET_WF"/>');
  results.push('<xcalcf:feature name="microsoft.com:LAMBDA_WF"/>');
  results.push('<xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/>');
  results.push("</xcalcf:calcFeatures>");
  results.push("</ext>");
  results.push("</extLst>");
  results.push("</workbook>");
  return results.join("");
}

function makeAppXml() {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
  );
  results.push("<Application>xlsx-kaku</Application>");
  results.push("<Manager></Manager>");
  results.push("<Company></Company>");
  results.push("<HyperlinkBase></HyperlinkBase>");
  results.push("<AppVersion>16.0300</AppVersion>");
  results.push("</Properties>");
  return results.join("");
}

function makeRelsFile() {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  );
  results.push(
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
  );
  results.push(
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
  );
  results.push(
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
  );
  results.push("</Relationships>");
  return results.join("");
}

function makeContentTypesXml(sharedStrings: boolean, sheetsLength: number) {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
  );
  results.push(
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  );
  results.push('<Default Extension="xml" ContentType="application/xml"/>');
  results.push(
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
  );
  results.push(
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
  );
  results.push(
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
  );

  // <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  // <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  for (let i = 1; i <= sheetsLength; i++) {
    results.push(
      `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
  }

  results.push(
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
  );
  results.push(
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  );
  if (sharedStrings) {
    results.push(
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    );
  }
  results.push("</Types>");
  return results.join("");
}
