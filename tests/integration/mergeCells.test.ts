import { readFileSync, rmSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import {
  deletePropertyFromObject,
  listFiles,
  parseXml,
  removeBasePath,
  unzip,
  writeFile,
} from "../helper/helper";
import { Workbook } from "../../src";

describe("mergeCells", () => {
  const testName = "mergeCells";

  const xlsxDir = "tests/xlsx";
  const outputDir = `tests/temp/${testName}/output`;

  const expectedUnzippedDir = `tests/temp/${testName}/expected`;
  const actualUnzippedDir = `tests/temp/${testName}/actual`;

  const expectedXlsxPath = resolve(xlsxDir, `${testName}.xlsx`);
  const actualXlsxPath = resolve(outputDir, `${testName}.xlsx`);

  const extension = extname(expectedXlsxPath);
  const xlsxBaseName = basename(expectedXlsxPath, extension);

  const expectedFileDir = resolve(expectedUnzippedDir, xlsxBaseName);
  const actualFileDir = resolve(actualUnzippedDir, xlsxBaseName);

  beforeAll(async () => {
    await unzip(expectedXlsxPath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 1 });
    ws.setCell(1, 0, { type: "number", value: 2 });
    ws.setMergeCell({ ref: "A1:C1" });
    ws.setMergeCell({ ref: "A2:A4" });

    const xlsx = await wb.generateXlsx();
    writeFile(actualXlsxPath, xlsx);

    await unzip(actualXlsxPath, actualFileDir);
  });

  afterAll(() => {
    rmSync(outputDir, { recursive: true });
    rmSync(expectedUnzippedDir, { recursive: true });
    rmSync(actualUnzippedDir, { recursive: true });
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

  test("app.xml", async () => {
    const expected = readFileSync(
      resolve(expectedFileDir, "docProps/app.xml"),
      "utf-8"
    );
    const actual = readFileSync(
      resolve(actualFileDir, "docProps/app.xml"),
      "utf-8"
    );

    const expectedObj = parseXml(expected);
    const actualObj = parseXml(actual);

    deletePropertyFromObject(expectedObj, "Properties.Application");
    deletePropertyFromObject(actualObj, "Properties.Application");

    expect(actualObj).toEqual(expectedObj);
  });

  test("core.xml", async () => {
    const expected = readFileSync(
      resolve(expectedFileDir, "docProps/core.xml"),
      "utf-8"
    );
    const actual = readFileSync(
      resolve(actualFileDir, "docProps/core.xml"),
      "utf-8"
    );

    const expectedObj = parseXml(expected);
    const actualObj = parseXml(actual);

    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "cp:coreProperties.dcterms:created");
    deletePropertyFromObject(actualObj, "cp:coreProperties.dcterms:created");

    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "cp:coreProperties.dcterms:modified");
    deletePropertyFromObject(actualObj, "cp:coreProperties.dcterms:modified");

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

    // It may be a problem-free difference.
    deletePropertyFromObject(expectedObj, "workbook.calcPr");

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
