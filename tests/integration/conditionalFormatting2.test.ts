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

describe("above average conditional formatting", () => {
  const testName = "conditionalFormatting2";

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
    ws.setCell(2, 0, { type: "number", value: 3 });
    ws.setCell(3, 0, { type: "number", value: 4 });
    ws.setCell(4, 0, { type: "number", value: 5 });
    ws.setCell(5, 0, { type: "number", value: 6 });
    ws.setCell(6, 0, { type: "number", value: 7 });
    ws.setCell(7, 0, { type: "number", value: 8 });
    ws.setCell(8, 0, { type: "number", value: 9 });
    ws.setCell(9, 0, { type: "number", value: 10 });
    ws.setCell(10, 0, { type: "number", value: 11 });
    ws.setCell(11, 0, { type: "number", value: 12 });
    ws.setCell(12, 0, { type: "number", value: 13 });
    ws.setCell(13, 0, { type: "number", value: 14 });
    ws.setCell(14, 0, { type: "number", value: 15 });
    ws.setCell(15, 0, { type: "number", value: 16 });
    ws.setCell(16, 0, { type: "number", value: 17 });
    ws.setCell(17, 0, { type: "number", value: 18 });
    ws.setCell(18, 0, { type: "number", value: 19 });
    ws.setCell(19, 0, { type: "number", value: 20 });

    ws.setConditionalFormatting({
      // Fixme: "A:A" is not supported
      sqref: "A1:A1048576",
      type: "aboveAverage",
      priority: 4,
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    ws.setCell(0, 1, { type: "number", value: 1 });
    ws.setCell(1, 1, { type: "number", value: 2 });
    ws.setCell(2, 1, { type: "number", value: 3 });
    ws.setCell(3, 1, { type: "number", value: 4 });
    ws.setCell(4, 1, { type: "number", value: 5 });
    ws.setCell(5, 1, { type: "number", value: 6 });
    ws.setCell(6, 1, { type: "number", value: 7 });
    ws.setCell(7, 1, { type: "number", value: 8 });
    ws.setCell(8, 1, { type: "number", value: 9 });
    ws.setCell(9, 1, { type: "number", value: 10 });
    ws.setCell(10, 1, { type: "number", value: 11 });
    ws.setCell(11, 1, { type: "number", value: 12 });
    ws.setCell(12, 1, { type: "number", value: 13 });
    ws.setCell(13, 1, { type: "number", value: 14 });
    ws.setCell(14, 1, { type: "number", value: 15 });
    ws.setCell(15, 1, { type: "number", value: 16 });
    ws.setCell(16, 1, { type: "number", value: 17 });
    ws.setCell(17, 1, { type: "number", value: 18 });
    ws.setCell(18, 1, { type: "number", value: 19 });
    ws.setCell(19, 1, { type: "number", value: 20 });

    ws.setConditionalFormatting({
      sqref: "B1:B1048576",
      type: "belowAverage",
      priority: 3,
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    ws.setCell(0, 2, { type: "number", value: 1 });
    ws.setCell(1, 2, { type: "number", value: 2 });
    ws.setCell(2, 2, { type: "number", value: 3 });
    ws.setCell(3, 2, { type: "number", value: 4 });
    ws.setCell(4, 2, { type: "number", value: 5 });
    ws.setCell(5, 2, { type: "number", value: 6 });
    ws.setCell(6, 2, { type: "number", value: 7 });
    ws.setCell(7, 2, { type: "number", value: 8 });
    ws.setCell(8, 2, { type: "number", value: 9 });
    ws.setCell(9, 2, { type: "number", value: 10 });
    ws.setCell(10, 2, { type: "number", value: 11 });
    ws.setCell(11, 2, { type: "number", value: 12 });
    ws.setCell(12, 2, { type: "number", value: 13 });
    ws.setCell(13, 2, { type: "number", value: 14 });
    ws.setCell(14, 2, { type: "number", value: 15 });
    ws.setCell(15, 2, { type: "number", value: 16 });
    ws.setCell(16, 2, { type: "number", value: 17 });
    ws.setCell(17, 2, { type: "number", value: 18 });
    ws.setCell(18, 2, { type: "number", value: 19 });
    ws.setCell(19, 2, { type: "number", value: 20 });

    ws.setConditionalFormatting({
      sqref: "C1:C1048576",
      type: "atOrAboveAverage",
      priority: 2,
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    ws.setCell(0, 3, { type: "number", value: 1 });
    ws.setCell(1, 3, { type: "number", value: 2 });
    ws.setCell(2, 3, { type: "number", value: 3 });
    ws.setCell(3, 3, { type: "number", value: 4 });
    ws.setCell(4, 3, { type: "number", value: 5 });
    ws.setCell(5, 3, { type: "number", value: 6 });
    ws.setCell(6, 3, { type: "number", value: 7 });
    ws.setCell(7, 3, { type: "number", value: 8 });
    ws.setCell(8, 3, { type: "number", value: 9 });
    ws.setCell(9, 3, { type: "number", value: 10 });
    ws.setCell(10, 3, { type: "number", value: 11 });
    ws.setCell(11, 3, { type: "number", value: 12 });
    ws.setCell(12, 3, { type: "number", value: 13 });
    ws.setCell(13, 3, { type: "number", value: 14 });
    ws.setCell(14, 3, { type: "number", value: 15 });
    ws.setCell(15, 3, { type: "number", value: 16 });
    ws.setCell(16, 3, { type: "number", value: 17 });
    ws.setCell(17, 3, { type: "number", value: 18 });
    ws.setCell(18, 3, { type: "number", value: 19 });
    ws.setCell(19, 3, { type: "number", value: 20 });

    ws.setConditionalFormatting({
      sqref: "D1:D1048576",
      type: "atOrBelowAverage",
      priority: 1,
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

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

    // Differences due to the default font
    deletePropertyFromObject(expectedObj, "styleSheet.fonts");
    // Differences due to the default font
    deletePropertyFromObject(actualObj, "styleSheet.fonts");

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

  test("WorkbookXmlRels", () => {
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

    // It may be a problem-free difference.
    deletePropertyFromObject(expectedObj, "worksheet.dimension.@_ref");
    deletePropertyFromObject(actualObj, "worksheet.dimension.@_ref");

    // It should be a problem-free difference.
    deletePropertyFromObject(
      expectedObj,
      "worksheet.sheetViews.sheetView.selection"
    );
    deletePropertyFromObject(
      actualObj,
      "worksheet.sheetViews.sheetView.selection"
    );

    // It should be a problem-free difference.
    // In oneline Excel, the ID of the last created element becomes 0.
    for (const c of expectedObj.worksheet.conditionalFormatting) {
      deletePropertyFromObject(c, "cfRule.@_dxfId");
    }
    // In xlsx-kaku, the ID of the first created element becomes 0.
    for (const c of actualObj.worksheet.conditionalFormatting) {
      deletePropertyFromObject(c, "cfRule.@_dxfId");
    }

    expect(actualObj).toEqual(expectedObj);
  });
});
