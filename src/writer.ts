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
import { strToU8, zip, zipSync } from "fflate";
import { makeWorksheetXml } from "./xml/worksheetXml";
import { makeAppXml } from "./xml/appXml";
import { makeSharedStringsXml } from "./xml/sharedStringsXml";
import { makeWorkbookXmlRels } from "./xml/workbookXmlRels";
import { makeCoreXml } from "./xml/coreXml";
import { makeStylesXml } from "./xml/stylesXml";
import { makeWorkbookXml } from "./xml/workbookXml";
import { makeRelsFile } from "./xml/rels";
import { makeContentTypesXml } from "./xml/contentTypesXml";
import { Dxf } from "./dxf";
import { DrawingRels } from "./drawingRels";
import { makeDrawingXml } from "./xml/drawingXml";

type CompressibleFile = {
  filename: string;
  content: string | Uint8Array;
};

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
  return compressXMLs(files);
}

export function genXlsxSync(worksheets: Worksheet[]) {
  const files = generateXMLs(worksheets);
  return compressXMLsSync(files);
}

function compressXMLs(files: CompressibleFile[]) {
  return new Promise<Uint8Array>((resolve, reject) => {
    const data: { [key: string]: Uint8Array } = {};

    for (const file of files) {
      if (file.content instanceof Uint8Array) {
        data[file.filename] = file.content;
      } else {
        data[file.filename] = strToU8(file.content);
      }
    }

    zip(data, (err, data) => {
      if (err) {
        reject(err);
        return;
      }

      resolve(data);
    });
  });
}

function compressXMLsSync(files: CompressibleFile[]) {
  const data: { [key: string]: Uint8Array } = {};

  for (const file of files) {
    if (file.content instanceof Uint8Array) {
      data[file.filename] = file.content;
    } else {
      data[file.filename] = strToU8(file.content);
    }
  }

  return zipSync(data);
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
    drawingRelsList,
    drawingXmlList,
  } = createExcelFiles(worksheets);

  const files: CompressibleFile[] = [];
  files.push({ filename: "[Content_Types].xml", content: contentTypesXml });
  files.push({ filename: "_rels/.rels", content: relsFile });
  files.push({ filename: "docProps/app.xml", content: appXml });
  files.push({ filename: "docProps/core.xml", content: coreXml });

  if (sharedStringsXml !== null) {
    files.push({
      filename: "xl/sharedStrings.xml",
      content: sharedStringsXml,
    });
  }

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

  for (let i = 0; i < drawingRelsList.length; i++) {
    files.push({
      filename: `xl/drawings/_rels/drawing${i + 1}.xml.rels`,
      content: drawingRelsList[i]!,
    });
  }

  for (let i = 0; i < drawingXmlList.length; i++) {
    files.push({
      filename: `xl/drawings/drawing${i + 1}.xml`,
      content: drawingXmlList[i]!,
    });
  }

  const images = worksheets.flatMap((worksheet) => worksheet.images);
  for (let i = 0; i < images.length; i++) {
    files.push({
      filename: `xl/media/image${i + 1}.${images[i]!.extension}`,
      content: images[i]!.data,
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

  const dxf = new Dxf();
  const drawingRels = new DrawingRels();

  const sheetXmlList: string[] = [];
  const worksheetRelsList: string[] = [];
  const worksheetsLength = worksheets.length;
  const drawingRelsList: string[] = [];
  const drawingXmlList: string[] = [];

  let count = 0;
  for (const worksheet of worksheets) {
    const { sheetXml, worksheetRels, drawingRelsXml, xlsxImages } =
      makeWorksheetXml(worksheet, styleMappers, dxf, drawingRels, count);

    sheetXmlList.push(sheetXml);
    if (worksheetRels !== null) {
      worksheetRelsList.push(worksheetRels);
    }

    if (drawingRelsXml !== null) {
      drawingRelsList.push(drawingRelsXml);
    }

    if (xlsxImages.length > 0) {
      const drawingXml = makeDrawingXml(xlsxImages);
      drawingXmlList.push(drawingXml);
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

  const imageExtensions = Array.from(
    new Set(
      worksheets.flatMap((worksheet) =>
        worksheet.images.map((image) => image.extension)
      )
    )
  );
  const contentTypesXml = makeContentTypesXml(
    imageExtensions,
    hasSharedStrings,
    worksheetsLength,
    drawingXmlList.length
  );

  const stylesXml = makeStylesXml(styleMappers, dxf);
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
    drawingRelsList,
    drawingXmlList,
  };
}
