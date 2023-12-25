import path, { basename, extname } from "node:path";
import {
  deletePropertyFromObject,
  listFiles,
  removeBasePath,
  unzip,
} from "./helper/helper";
import { Workbook } from "../src/index";
import { readFileSync, rmSync } from "node:fs";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

const OUTPUT_DIR = "tests/output";
const XLSX_Dir = "tests/xlsx";
const EXPECTED_UNZIPPED_DIR = "tests/expected";
const ACTUAL_UNZIPPED_DIR = "tests/actuall";

const parser = new XMLParser({ ignoreAttributes: false });
const builder = new XMLBuilder();

describe("number", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = path.resolve(XLSX_Dir, "number.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = path.resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 1 });
    ws.setCell(0, 1, { type: "number", value: 2 });
    ws.setCell(1, 0, { type: "number", value: 3 });
    ws.setCell(1, 1, { type: "number", value: 4 });
    outputPath = path.resolve(OUTPUT_DIR, "number.xlsx");
    await wb.save(outputPath);

    actualFileDir = path.resolve(ACTUAL_UNZIPPED_DIR, xlsxBaseName);
    await unzip(outputPath, actualFileDir);
  });

  afterAll(() => {
    rmSync(OUTPUT_DIR, { recursive: true });
    rmSync(EXPECTED_UNZIPPED_DIR, { recursive: true });
    rmSync(ACTUAL_UNZIPPED_DIR, { recursive: true });
  });

  test("compare files", async () => {
    const expectedFiles = listFiles(expectedFileDir);
    const actualFiles = listFiles(actualFileDir);

    const expectedSubPaths = expectedFiles.map((it) =>
      removeBasePath(it, expectedFileDir)
    );
    const actualSubPaths = actualFiles.map((it) =>
      removeBasePath(it, actualFileDir)
    );

    expect(actualSubPaths).toEqual(expectedSubPaths);
  });

  test("compare StylesXml", () => {
    const expectedXlStylesXmlPath = path.resolve(
      expectedFileDir,
      "xl/styles.xml"
    );
    const expectedXlStylesXml = readFileSync(expectedXlStylesXmlPath, "utf8");

    const expectedObj = parser.parse(expectedXlStylesXml);
    // Differences due to the default font
    deletePropertyFromObject(expectedObj, "styleSheet.fonts");
    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "styleSheet.dxfs");
    const expectedXlStylesXmlWithoutFonts = builder.build(
      expectedObj.styleSheet
    );

    const actualXlStylesXmlPath = path.resolve(actualFileDir, "xl/styles.xml");
    const actualXlStylesXml = readFileSync(actualXlStylesXmlPath, "utf8");
    const actualObj = parser.parse(actualXlStylesXml);
    // Differences due to the default font
    deletePropertyFromObject(actualObj, "styleSheet.fonts");
    // It should be a problem-free difference.
    deletePropertyFromObject(actualObj, "styleSheet.cellStyleXfs.xf.@_xfId");
    const actualXlStylesXmlWithoutFonts = builder.build(actualObj.styleSheet);

    expect(actualXlStylesXmlWithoutFonts).toEqual(
      expectedXlStylesXmlWithoutFonts
    );
  });

  test("compare WorkbookXmlRels", () => {
    function sortById(a: any, b: any) {
      const rIdA = parseInt(a["@_Id"].substring(3));
      const rIdB = parseInt(b["@_Id"].substring(3));
      if (rIdA < rIdB) {
        return -1;
      }
      if (rIdA > rIdB) {
        return 1;
      }
      return 0;
    }

    const expectedXlWorkbookXmlRelsPath = path.resolve(
      expectedFileDir,
      "xl/_rels/workbook.xml.rels"
    );
    const expectedXlWorkbookXmlRels = readFileSync(
      expectedXlWorkbookXmlRelsPath,
      "utf8"
    );
    const actualXlWorkbookXmlRelsPath = path.resolve(
      actualFileDir,
      "xl/_rels/workbook.xml.rels"
    );
    const actualXlWorkbookXmlRels = readFileSync(
      actualXlWorkbookXmlRelsPath,
      "utf8"
    );

    const expectedObj = parser.parse(expectedXlWorkbookXmlRels);
    const actualObj = parser.parse(actualXlWorkbookXmlRels);

    const expectedRelationships = expectedObj.Relationships.Relationship;
    expectedRelationships.sort(sortById);

    const actualRelationships = actualObj.Relationships.Relationship;
    actualRelationships.sort(sortById);

    expect(actualRelationships).toEqual(expectedRelationships);
  });

  test("compare WorkbookXml", () => {
    const expectedXlWorkbookXmlPath = path.resolve(
      expectedFileDir,
      "xl/workbook.xml"
    );
    const expectedXlWorkbookXml = readFileSync(
      expectedXlWorkbookXmlPath,
      "utf8"
    );
    const actualXlWorkbookXmlPath = path.resolve(
      actualFileDir,
      "xl/workbook.xml"
    );
    const actualXlWorkbookXml = readFileSync(actualXlWorkbookXmlPath, "utf8");

    const expectedObj = parser.parse(expectedXlWorkbookXml);
    const actualObj = parser.parse(actualXlWorkbookXml);

    // It should be a problem-free difference.
    deletePropertyFromObject(
      expectedObj,
      "workbook.xr:revisionPtr.@_documentId"
    );
    deletePropertyFromObject(actualObj, "workbook.xr:revisionPtr.@_documentId");

    expect(actualObj).toEqual(expectedObj);
  });

  test("compare worksheets", () => {
    const expectedXlSheet1XmlPath = path.resolve(
      expectedFileDir,
      "xl/worksheets/sheet1.xml"
    );
    const expectedXlSheet1Xml = readFileSync(expectedXlSheet1XmlPath, "utf8");
    const actualXlSheet1XmlPath = path.resolve(
      actualFileDir,
      "xl/worksheets/sheet1.xml"
    );
    const actualXlSheet1Xml = readFileSync(actualXlSheet1XmlPath, "utf8");

    const expectedObj = parser.parse(expectedXlSheet1Xml);
    const actualObj = parser.parse(actualXlSheet1Xml);

    // It should be a problem-free difference.
    deletePropertyFromObject(
      expectedObj,
      "worksheet.sheetViews.sheetView.selection"
    );
    deletePropertyFromObject(
      actualObj,
      "worksheet.sheetViews.sheetView.selection"
    );

    expect(actualObj).toEqual(expectedObj);
  });
});
