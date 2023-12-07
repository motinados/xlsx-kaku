import * as fs from "node:fs";
import path from "node:path";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import { Cell, NullableCell, convNumberToColumn } from "./sheetData";
import { SharedStrings } from "./sharedStrings";
import { makeThemeXml } from "./theme";
import { Fills } from "./fills";
import { CellXf, CellXfs } from "./cellXfs";
import { Fonts } from "./fonts";
import { Borders } from "./borders";
import { NumberFormats } from "./numberFormats";
import { CellStyles } from "./cellStyles";
import { CellStyleXfs } from "./cellStyleXfs";
import { Hyperlinks } from "./hyperlinks";
import { WorksheetRels } from "./worksheetRels";

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
};

export async function writeFile(filename: string, sheetData: NullableCell[][]) {
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

  const { sheetDataXml, sharedStringsXml } = tableToString(
    sheetData,
    styleMappers
  );
  const hasSharedStrings = sharedStringsXml !== null;
  const dimension = getDimension(sheetData);
  const sheetXml = makeSheetXml(
    sheetDataXml,
    dimension,
    styleMappers.hyperlinks
  );
  const themeXml = makeThemeXml();
  const appXml = makeAppXml();
  const coreXml = makeCoreXml();
  const stylesXml = makeStylesXml(styleMappers);
  const workbookXml = makeWorkbookXml();
  const workbookXmlRels = makeWorkbookXmlRels(hasSharedStrings);
  const relsFile = makeRelsFile();
  const contentTypesFile = makeContentTypesFile(hasSharedStrings);

  const xlsxPath = path.resolve(filename);
  if (!fs.existsSync(xlsxPath)) {
    fs.mkdirSync(xlsxPath, { recursive: true });
  }
  fs.writeFileSync(
    path.join(xlsxPath, "[Content_Types].xml"),
    contentTypesFile
  );

  const _relsPath = path.resolve(xlsxPath, "_rels");
  if (!fs.existsSync(_relsPath)) {
    fs.mkdirSync(_relsPath, { recursive: true });
  }
  fs.writeFileSync(path.join(_relsPath, ".rels"), relsFile);

  const docPropsPath = path.resolve(xlsxPath, "docProps");
  if (!fs.existsSync(docPropsPath)) {
    fs.mkdirSync(docPropsPath, { recursive: true });
  }
  fs.writeFileSync(path.join(docPropsPath, "app.xml"), appXml);
  fs.writeFileSync(path.join(docPropsPath, "core.xml"), coreXml);

  const xlPath = path.resolve(xlsxPath, "xl");
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
  fs.writeFileSync(path.join(worksheetsPath, "sheet1.xml"), sheetXml);

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

  await zipToXlsx(xlsxPath, xlsxPath + ".xlsx");
}

export function zipToXlsx(sourceDir: string, outPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => {
      console.log(`Archived ${archive.pointer()} total bytes`);
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

export function findFirstNonNullCell(row: NullableCell[]) {
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
export function findLastNonNullCell(row: NullableCell[]) {
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

export function tableToString(
  table: NullableCell[][],
  styleMappers: StyleMappers
) {
  const sheetDataXml = makeSheetDataXml(table, styleMappers);
  const sharedStringsXml = makeSharedStringsXml(styleMappers.sharedStrings);
  return { sheetDataXml, sharedStringsXml };
}

export function makeSheetXml(
  sheetDataString: string,
  dimension: { start: string; end: string },
  hyperlinks: Hyperlinks
) {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">'
  );
  results.push(`<dimension ref="${dimension.start}:${dimension.end}"/>`);
  results.push("<sheetViews>");
  results.push(
    `<sheetView tabSelected="1" workbookViewId="0"><selection activeCell="${dimension.start}" sqref="${dimension.start}"/></sheetView>`
  );
  results.push("</sheetViews>");
  results.push('<sheetFormatPr defaultRowHeight="13.5"/>');
  results.push(sheetDataString);

  if (hyperlinks.getHyperlinks().length > 0) {
    results.push(hyperlinks.makeXML());
  }

  results.push(
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
  );
  return results.join("");
}

export function makeSharedStringsXml(sharedStrings: SharedStrings) {
  if (sharedStrings.count === 0) {
    return null;
  }

  let result = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.count}" uniqueCount="${sharedStrings.uniqueCount}">`;
  for (const str of sharedStrings.getValuesInOrder()) {
    result += `<si><t>${str}</t></si>`;
  }
  result += `</sst>`;
  return result;
}

function findFirstNotBlankRow(table: NullableCell[][]) {
  let index = 0;
  for (let i = 0; i < table.length; i++) {
    const row = table[i]!;
    if (row.length > 0) {
      index = i;
      break;
    }
  }
  return index;
}

function findLastNotBlankRow(table: NullableCell[][]) {
  let index = 0;
  for (let i = table.length - 1; i >= 0; i--) {
    const row = table[i]!;
    if (row.length > 0) {
      index = i;
      break;
    }
  }
  return index;
}

export function getDimension(sheetData: NullableCell[][]) {
  const firstRowIndex = findFirstNotBlankRow(sheetData);
  const lastRowIndex = findLastNotBlankRow(sheetData);
  if (firstRowIndex === null || lastRowIndex === null) {
    throw new Error("sheetData is empty");
  }

  const firstRowNumber = firstRowIndex + 1;
  const lastRowNumber = lastRowIndex + 1;

  const spans = getSpansFromTable(sheetData);
  const { startNumber, endNumber } = spans;
  const firstColumn = convNumberToColumn(startNumber - 1);
  const lastColumn = convNumberToColumn(endNumber - 1);

  return {
    start: `${firstColumn}${firstRowNumber}`,
    end: `${lastColumn}${lastRowNumber}`,
  };
}

export function makeSheetDataXml(
  table: NullableCell[][],
  styleMappers: StyleMappers
) {
  const { startNumber, endNumber } = getSpansFromTable(table);

  let result = `<sheetData>`;
  let rowIndex = 0;
  for (const row of table) {
    const str = rowToString(
      row,
      rowIndex,
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
  row: NullableCell[],
  rowIndex: number,
  startNumber: number,
  endNumber: number,
  styleMappers: StyleMappers
): string | null {
  if (row.length === 0) {
    return null;
  }

  const rowNumber = rowIndex + 1;
  let result = `<row r="${rowNumber}" spans="${startNumber}:${endNumber}">`;

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

export function getSpansFromTable(table: NullableCell[][]) {
  const all = table
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

export function getSpans(row: NullableCell[]) {
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

function getCellXfId(xlsxCellStyle: XlsxCellStyle | null, cellXfs: CellXfs) {
  if (xlsxCellStyle) {
    const style = {
      fillId: xlsxCellStyle.fillId || 0,
      fontId: xlsxCellStyle.fontId || 0,
      borderId: xlsxCellStyle.borderId || 0,
      numFmtId: xlsxCellStyle.numFmtId || 0,
    };
    const xfId = cellXfs.getCellXfId(style);
    return xfId;
  }

  return null;
}

function getCellStyleXfId(
  xlsxCellStyle: XlsxCellStyle | null,
  cellStyleXfs: CellStyleXfs
) {
  if (xlsxCellStyle) {
    const style = {
      fillId: xlsxCellStyle.fillId || 0,
      fontId: xlsxCellStyle.fontId || 0,
      borderId: xlsxCellStyle.borderId || 0,
      numFmtId: xlsxCellStyle.numFmtId || 0,
    };
    const xfId = cellStyleXfs.getCellStyleXfId(style);
    return xfId;
  }

  return null;
}

export function getXlsxCellStyle(
  cell: Cell,
  mappers: StyleMappers
): XlsxCellStyle | null {
  if (cell.style) {
    const style = {
      fillId: cell.style.fill ? mappers.fills.getFillId(cell.style.fill) : 0,
      fontId: cell.style.font ? mappers.fonts.getFontId(cell.style.font) : 0,
      borderId: cell.style.border
        ? mappers.borders.getBorderId(cell.style.border)
        : 0,
      numFmtId: cell.style.numberFormat
        ? mappers.numberFormats.getNumFmtId(cell.style.numberFormat.formatCode)
        : 0,
    };
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
  const xlsxCellStyle = getXlsxCellStyle(cell, styleMappers);

  switch (cell.type) {
    case "number": {
      const cellXfId = getCellXfId(xlsxCellStyle, styleMappers.cellXfs);
      const s = cellXfId !== null ? ` s="${cellXfId}"` : "";
      return `<c r="${column}${rowNumber}"${s}><v>${cell.value}</v></c>`;
    }
    case "string": {
      const cellXfId = getCellXfId(xlsxCellStyle, styleMappers.cellXfs);
      const s = cellXfId !== null ? ` s="${cellXfId}"` : "";
      const index = styleMappers.sharedStrings.getIndex(cell.value);
      return `<c r="${column}${rowNumber}"${s} t="s"><v>${index}</v></c>`;
    }
    case "date": {
      const cellXfId = getCellXfId(xlsxCellStyle, styleMappers.cellXfs);
      const s = cellXfId !== null ? ` s="${cellXfId}"` : "";
      const serialValue = convertIsoStringToSerialValue(cell.value);
      return `<c r="${column}${rowNumber}"${s}><v>${serialValue}</v></c>`;
    }
    case "hyperlink": {
      if (xlsxCellStyle === null) {
        throw new Error("xlsxCellStyle is null for hyperlink");
      }
      const xfId = getCellStyleXfId(xlsxCellStyle, styleMappers.cellStyleXfs);
      if (xfId === null) {
        throw new Error("xfId is null for hyperlink");
      }

      const cellXf: CellXf = {
        xfId: xfId,
        ...xlsxCellStyle,
      };
      const cellXfId = styleMappers.cellXfs.getCellXfId(cellXf);
      const s = cellXfId !== null ? ` s="${cellXfId}"` : "";
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
    // default: {
    //   throw new Error(`not implemented: ${cell.type}`);
    // }
  }
}

function makeWorkbookXmlRels(sharedStrings: boolean): string {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  );
  results.push(
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
  );
  results.push(
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
  );
  results.push(
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
  );
  if (sharedStrings) {
    results.push(
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
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

function makeWorkbookXml() {
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
  results.push('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>');
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
  results.push("<Application>excel-writer</Application>");
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

function makeContentTypesFile(sharedStrings: boolean) {
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
  results.push(
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
  );
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
