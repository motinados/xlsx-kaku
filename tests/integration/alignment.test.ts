import { readFileSync, rmSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import {
  deletePropertyFromObject,
  listFiles,
  parseXml,
  removeBasePath,
  unzip,
} from "../helper/helper";
import { Workbook } from "../../src";

describe("alignment", () => {
  const XLSX_Dir = "tests/xlsx";
  const OUTPUT_DIR = "tests/temp/alignment/output";
  const EXPECTED_UNZIPPED_DIR = "tests/temp/alignment/expected";
  const ACTUAL_UNZIPPED_DIR = "tests/temp/alignment/actuall";

  const filepath = resolve(XLSX_Dir, "alignment.xlsx");
  const extension = extname(filepath);
  const xlsxBaseName = basename(filepath, extension);
  const expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
  const outputPath = resolve(OUTPUT_DIR, "alignment.xlsx");
  const actualFileDir = resolve(ACTUAL_UNZIPPED_DIR, xlsxBaseName);

  beforeAll(async () => {
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.setCell(0, 0, {
      type: "number",
      value: 1,
      style: { alignment: { vertical: "top" } },
    });
    ws.setCell(0, 1, {
      type: "number",
      value: 2,
      style: { alignment: { vertical: "center" } },
    });
    ws.setCell(0, 2, {
      type: "number",
      value: 3,
      style: { alignment: { vertical: "bottom" } },
    });
    ws.setCell(1, 0, {
      type: "number",
      value: 4,
      style: { alignment: { horizontal: "left" } },
    });
    ws.setCell(1, 1, {
      type: "number",
      value: 5,
      style: {
        alignment: {
          horizontal: "center",
        },
      },
    });
    ws.setCell(1, 2, {
      type: "number",
      value: 6,
      style: { alignment: { horizontal: "right" } },
    });
    ws.setCell(2, 0, {
      type: "number",
      value: 7,
      style: { alignment: { textRotation: 45 } },
    });
    ws.setCell(2, 1, {
      type: "number",
      value: 8,
      style: { alignment: { textRotation: 135 } },
    });
    ws.setCell(2, 2, {
      type: "number",
      value: 9,
      style: { alignment: { textRotation: 90 } },
    });
    ws.setCell(3, 0, {
      type: "number",
      value: 10,
      style: { alignment: { horizontal: "left", textRotation: 90 } },
    });
    ws.setCell(3, 1, {
      type: "number",
      value: 11,
      style: { alignment: { horizontal: "center", textRotation: 180 } },
    });
    ws.setCell(3, 2, {
      type: "number",
      value: 12,
      style: {
        alignment: { horizontal: "center", vertical: "top", textRotation: 255 },
      },
    });

    ws.setRowHeight({ index: 0, height: 39.75 });
    ws.setRowHeight({ index: 1, height: 39.75 });
    ws.setRowHeight({ index: 2, height: 39.75 });
    ws.setRowHeight({ index: 3, height: 39.75 });

    await wb.save(outputPath);

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

    // <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>
    // Probably "bottom" is the default value, so it seems that it does not appear in the expected XML.
    const obj = actualObj.styleSheet.cellXfs.xf[3];
    deletePropertyFromObject(obj, "alignment");

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
