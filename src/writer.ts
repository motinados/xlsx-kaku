import { SharedStrings } from "./sharedStrings";
import { makeThemeXml } from "./xml/themeXml";
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
import { makeAppXml } from "./xml/appXml";
import { makeSharedStringsXml } from "./xml/sharedStringsXml";
import { makeWorkbookXmlRels } from "./xml/workbookXmlRels";
import { makeCoreXml } from "./xml/coreXml";
import { makeStylesXml } from "./xml/stylesXml";
import { makeWorkbookXml } from "./xml/workbookXml";
import { makeRelsFile } from "./xml/relsFile";
import { makeContentTypesXml } from "./xml/contentTypesXml";

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
    sheetXmlList,
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

  for (let i = 0; i < sheetXmlList.length; i++) {
    files.push({
      filename: `xl/worksheets/sheet${i + 1}.xml`,
      content: sheetXmlList[i]!,
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

  const sheetXmlList: string[] = [];
  const worksheetRelsList: string[] = [];
  const worksheetsLength = worksheets.length;

  let count = 0;
  for (const worksheet of worksheets) {
    const { sheetXml, worksheetRels } = makeWorksheetXml(
      worksheet,
      styleMappers,
      count
    );

    sheetXmlList.push(sheetXml);
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
    sheetXmlList,
    worksheetRelsList,
  };
}
