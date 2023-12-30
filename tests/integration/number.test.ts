import path, { basename, extname } from "node:path";
import {
  deletePropertyFromObject,
  listFiles,
  parseXml,
  removeBasePath,
  unzip,
} from "../helper/helper";
import { Workbook } from "../../src/index";
import { readFileSync, rmSync } from "node:fs";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/number/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/number/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/number/actuall";

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
    const expectedXmlPath = path.resolve(expectedFileDir, "xl/styles.xml");
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const expectedObj = parseXml(expectedXml);

    // Differences due to the default font
    deletePropertyFromObject(expectedObj, "styleSheet.fonts");
    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "styleSheet.dxfs");

    const actualXmlPath = path.resolve(actualFileDir, "xl/styles.xml");
    const actualXml = readFileSync(actualXmlPath, "utf8");
    const actualObj = parseXml(actualXml);

    // Differences due to the default font
    deletePropertyFromObject(actualObj, "styleSheet.fonts");
    // It should be a problem-free difference.
    deletePropertyFromObject(actualObj, "styleSheet.cellStyleXfs.xf.@_xfId");

    expect(actualObj).toEqual(expectedObj);
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

    const expectedRelsPath = path.resolve(
      expectedFileDir,
      "xl/_rels/workbook.xml.rels"
    );
    const expectedRels = readFileSync(expectedRelsPath, "utf8");
    const actualRelsPath = path.resolve(
      actualFileDir,
      "xl/_rels/workbook.xml.rels"
    );
    const actualRels = readFileSync(actualRelsPath, "utf8");

    const expectedObj = parseXml(expectedRels);
    const actualObj = parseXml(actualRels);

    const expectedRelationships = expectedObj.Relationships.Relationship;
    expectedRelationships.sort(sortById);

    const actualRelationships = actualObj.Relationships.Relationship;
    actualRelationships.sort(sortById);

    expect(actualRelationships).toEqual(expectedRelationships);
  });

  test("compare WorkbookXml", () => {
    const expectedXmlPath = path.resolve(expectedFileDir, "xl/workbook.xml");
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const actualXmlPath = path.resolve(actualFileDir, "xl/workbook.xml");
    const actualXml = readFileSync(actualXmlPath, "utf8");

    const expectedObj = parseXml(expectedXml);
    const actualObj = parseXml(actualXml);

    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "workbook.fileVersion.@_rupBuild");
    deletePropertyFromObject(actualObj, "workbook.fileVersion.@_rupBuild");

    // It should be a problem-free difference.
    deletePropertyFromObject(
      expectedObj,
      "workbook.xr:revisionPtr.@_documentId"
    );
    deletePropertyFromObject(actualObj, "workbook.xr:revisionPtr.@_documentId");

    expect(actualObj).toEqual(expectedObj);
  });

  test("compare worksheets", () => {
    const expectedXmlPath = path.resolve(
      expectedFileDir,
      "xl/worksheets/sheet1.xml"
    );
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const actualXmlPath = path.resolve(
      actualFileDir,
      "xl/worksheets/sheet1.xml"
    );
    const actualXml = readFileSync(actualXmlPath, "utf8");

    const expectedObj = parseXml(expectedXml);
    const actualObj = parseXml(actualXml);

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
