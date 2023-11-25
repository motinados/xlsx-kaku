import * as fs from "node:fs";
import path from "node:path";

import { NullableCell, convNumberToColumn } from "./sheetData";
import { SharedStrings } from "./sharedStrings";

export function writeFile(filename: string) {
  const xlsxPath = path.resolve(filename);
  if (!fs.existsSync(xlsxPath)) {
    fs.mkdirSync(xlsxPath, { recursive: true });
  }

  const _relsPath = path.resolve(xlsxPath, "_rels");
  if (!fs.existsSync(_relsPath)) {
    fs.mkdirSync(_relsPath, { recursive: true });
  }

  const docPropsPath = path.resolve(xlsxPath, "docProps");
  if (!fs.existsSync(docPropsPath)) {
    fs.mkdirSync(docPropsPath, { recursive: true });
  }

  const coreXml = makeCoreXml();
  fs.writeFileSync(path.join(docPropsPath, "core.xml"), coreXml);

  const xlPath = path.resolve(xlsxPath, "xl");
  if (!fs.existsSync(xlPath)) {
    fs.mkdirSync(xlPath, { recursive: true });
  }

  const stylesXml = makeStylesXml();
  fs.writeFileSync(path.join(xlPath, "styles.xml"), stylesXml);
  const workbookXml = makeWorkbookXml();
  fs.writeFileSync(path.join(xlPath, "workbook.xml"), workbookXml);

  const xl_relsPath = path.resolve(xlPath, "_rels");
  if (!fs.existsSync(xl_relsPath)) {
    fs.mkdirSync(xl_relsPath, { recursive: true });
  }

  const workbookXmlRels = makeWorkbookXmlRels(true);
  fs.writeFileSync(
    path.join(xl_relsPath, "workbook.xml.rels"),
    workbookXmlRels
  );

  const themePath = path.resolve(xlPath, "theme");
  if (!fs.existsSync(themePath)) {
    fs.mkdirSync(themePath, { recursive: true });
  }

  const worksheetsPath = path.resolve(xlPath, "worksheets");
  if (!fs.existsSync(worksheetsPath)) {
    fs.mkdirSync(worksheetsPath, { recursive: true });
  }
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

export function tableToString(table: NullableCell[][]) {
  const sharedStrings = new SharedStrings();

  const sheetDataString = makeSheetDataXml(table, sharedStrings);
  const sharedStringsXml = makeSharedStringsXml(sharedStrings);
  return { sheetDataString, sharedStringsXml };
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
  sharedStrings: SharedStrings
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
      sharedStrings
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
  sharedStrings: SharedStrings
): string | null {
  if (row.length === 0) {
    return null;
  }

  const rowNumber = rowIndex + 1;
  let result = `<row r="${rowNumber}" spans="${startNumber}:${endNumber}">`;

  let columnIndex = 0;
  for (const cell of row) {
    if (cell !== null) {
      result += cellToString(cell, columnIndex, rowIndex, sharedStrings);
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

export function cellToString(
  cell: NonNullable<NullableCell>,
  columnIndex: number,
  rowIndex: number,
  sharedStrings: SharedStrings
) {
  const rowNumber = rowIndex + 1;
  const column = convNumberToColumn(columnIndex);
  switch (cell.type) {
    case "number": {
      return `<c r="${column}${rowNumber}"><v>${cell.value}</v></c>`;
    }
    case "string": {
      const index = sharedStrings.getIndex(cell.value);
      return `<c r="${column}${rowNumber}" t="s"><v>${index}</v></c>`;
    }
    default: {
      throw new Error(`not implemented: ${cell.type}`);
    }
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

function makeStylesXml() {
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">'
  );
  results.push('<fonts count="1">');
  results.push("<font>");
  results.push('<sz val="11"/>');
  results.push('<color theme="1"/>');
  results.push('<name val="Calibri"/>');
  results.push('<family val="2"/>');
  results.push('<scheme val="minor"/></font>');
  results.push("</fonts>");
  results.push('<fills count="2">');
  results.push('<fill><patternFill patternType="none"/></fill>');
  results.push('<fill><patternFill patternType="gray125"/></fill>');
  results.push("</fills>");
  results.push('<borders count="1">');
  results.push("<border><left/><right/><top/><bottom/><diagonal/></border>");
  results.push("</borders>");
  results.push(
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  );
  results.push(
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  );
  results.push(
    '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>'
  );
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
  const results: string[] = [];
  results.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  results.push(
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  );
  results.push(
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="27123"/>'
  );
  results.push('<workbookPr defaultThemeVersion="166925"/>');
  results.push("<bookViews>");
  results.push(
    '<workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010"/>'
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
