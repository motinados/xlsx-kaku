import { v4 as uuidv4 } from "uuid";
import { SharedStrings } from "./sharedStrings";
import { makeThemeXml } from "./theme";
import { Fills } from "./fills";
import { CellXfs } from "./cellXfs";
import { Fonts } from "./fonts";
import { Borders } from "./borders";
import { NumberFormats } from "./numberFormats";
import { CellStyles } from "./cellStyles";
import { CellStyleXfs } from "./cellStyleXfs";
import { Hyperlinks } from "./hyperlinks";
import { WorksheetRels } from "./worksheetRels";
import { Worksheet } from "./worksheet";
import { strToU8, zipSync } from "fflate";
import { makeWorksheetXml } from "./xml/worksheetXml";

export type StyleMappers = {
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
    worksheetRelsList,
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

  for (let i = 0; i < worksheetRelsList.length; i++) {
    files.push({
      filename: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
      content: worksheetRelsList[i]!,
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
  const worksheetRelsList: string[] = [];
  const worksheetsLength = worksheets.length;

  let count = 0;
  for (const worksheet of worksheets) {
    const { sheetXml, worksheetRels } = makeWorksheetXml(
      worksheet,
      styleMappers,
      count
    );

    sheetXmls.push(sheetXml);
    if (worksheetRels !== null) {
      worksheetRelsList.push(worksheetRels);
    }

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
    worksheetRelsList,
  };
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
