import { readFileSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import {
  deletePropertyFromObject,
  listFiles,
  parseXml,
  removeBasePath,
  unzip,
} from "../helper/helper";
import { Workbook } from "../../src";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/mergeCells/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/mergeCells/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/mergeCells/actuall";

describe("string", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = resolve(XLSX_Dir, "mergeCells.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 1 });
    ws.setCell(1, 0, { type: "number", value: 2 });
    ws.setMergeCell({ ref: "A1:C1" });
    ws.setMergeCell({ ref: "A2:A4" });

    outputPath = resolve(OUTPUT_DIR, "mergeCells.xlsx");
    await wb.save(outputPath);

    actualFileDir = resolve(ACTUAL_UNZIPPED_DIR, xlsxBaseName);
    await unzip(outputPath, actualFileDir);
  });

  afterAll(() => {
    // rmSync(OUTPUT_DIR, { recursive: true });
    // rmSync(EXPECTED_UNZIPPED_DIR, { recursive: true });
    // rmSync(ACTUAL_UNZIPPED_DIR, { recursive: true });
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

  test("Content_Types.xml", async () => {
    const expected = readFileSync(
      resolve(expectedFileDir, "[Content_Types].xml"),
      "utf-8"
    );
    const actual = readFileSync(
      resolve(actualFileDir, "[Content_Types].xml"),
      "utf-8"
    );

    const expectedObj = parseXml(expected);
    const actualObj = parseXml(actual);

    expect(actualObj).toEqual(expectedObj);
  });

  test("styles.xml", async () => {
    const expected = readFileSync(
      resolve(expectedFileDir, "xl/styles.xml"),
      "utf-8"
    );
    const actual = readFileSync(
      resolve(actualFileDir, "xl/styles.xml"),
      "utf-8"
    );

    const expectedObj = parseXml(expected);
    const actualObj = parseXml(actual);

    // // Differences due to the default font
    deletePropertyFromObject(expectedObj, "styleSheet.fonts");
    // // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "styleSheet.dxfs");
    // // Differences due to the default font
    deletePropertyFromObject(actualObj, "styleSheet.fonts");

    // In online Excel, when merging cells, styles are automatically added.
    // This is the difference caused by those styles.
    deletePropertyFromObject(expectedObj, "styleSheet.cellXfs");
    deletePropertyFromObject(actualObj, "styleSheet.cellXfs");

    expect(actualObj).toEqual(expectedObj);
  });

  test("workbookXml", () => {
    const expectedXmlPath = resolve(expectedFileDir, "xl/workbook.xml");
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const actualXmlPath = resolve(actualFileDir, "xl/workbook.xml");
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

  test("workbookXmlRels", () => {
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

    const expectedRelsPath = resolve(
      expectedFileDir,
      "xl/_rels/workbook.xml.rels"
    );
    const expectedRels = readFileSync(expectedRelsPath, "utf8");
    const actualRelsPath = resolve(actualFileDir, "xl/_rels/workbook.xml.rels");
    const actualRels = readFileSync(actualRelsPath, "utf8");

    const expectedObj = parseXml(expectedRels);
    const actualObj = parseXml(actualRels);

    const expectedRelationships = expectedObj.Relationships.Relationship;
    expectedRelationships.sort(sortById);

    const actualRelationships = actualObj.Relationships.Relationship;
    actualRelationships.sort(sortById);

    expect(actualRelationships).toEqual(expectedRelationships);
  });

  test("worksheets", () => {
    const expectedXmlPath = resolve(
      expectedFileDir,
      "xl/worksheets/sheet1.xml"
    );
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const actualXmlPath = resolve(actualFileDir, "xl/worksheets/sheet1.xml");
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

    // In online Excel, when merging cells, styles are automatically added.
    // This is the difference caused by those styles.
    for (const obj of expectedObj.worksheet.sheetData.row) {
      deletePropertyFromObject(obj, "@_ht");
      deletePropertyFromObject(obj, "@_customHeight");

      if (!Array.isArray(obj.c)) {
        deletePropertyFromObject(obj.c, "@_s");
      } else {
        for (const c of obj.c) {
          deletePropertyFromObject(c, "@_s");
        }
      }
    }

    expect(actualObj).toEqual(expectedObj);
  });
});
