// import { rmSync } from "node:fs";
import { readFileSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import { XMLParser } from "fast-xml-parser";
import {
  deletePropertyFromObject,
  listFiles,
  removeBasePath,
  unzip,
  writeFile,
} from "../helper/helper";
import { Workbook } from "../../src";

const parser = new XMLParser({ ignoreAttributes: false });

describe("date", () => {
  const testName = "date";

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
    ws.setCell(0, 0, {
      type: "date",
      value: "2023-01-02T00:00:00.000Z",
      style: {
        numberFormat: {
          formatCode: "yyyy-mm-dd",
        },
      },
    });
    ws.setCell(0, 1, {
      type: "date",
      value: "2023-01-02T00:00:00.000Z",
      style: {
        numberFormat: {
          formatCode: "m/d/yy",
        },
      },
    });
    ws.setCell(1, 0, {
      type: "date",
      value: "2023-01-03T00:00:00.000Z",
      style: {
        numberFormat: {
          formatCode: "yyyy-mm-dd",
        },
      },
    });
    ws.setCell(1, 1, {
      type: "date",
      value: "2023-01-03T00:00:00.000Z",
      style: {
        numberFormat: {
          formatCode: "m/d/yy",
        },
      },
    });
    const xlsx = wb.generateXlsx();
    writeFile(actualXlsxPath, xlsx);
    await unzip(actualXlsxPath, actualFileDir);
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

    const expectedObj = parser.parse(expected);
    const actualObj = parser.parse(actual);

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

    const expectedObj = parser.parse(expected);
    const actualObj = parser.parse(actual);

    // Differences due to the default font
    deletePropertyFromObject(expectedObj, "styleSheet.fonts");
    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "styleSheet.dxfs");
    // Differences due to the default font
    deletePropertyFromObject(actualObj, "styleSheet.fonts");
    // It should be a problem-free difference.
    deletePropertyFromObject(actualObj, "styleSheet.cellStyleXfs.xf.@_xfId");

    expect(actualObj).toEqual(expectedObj);
  });

  test("workbookXml", () => {
    const expectedXmlPath = resolve(expectedFileDir, "xl/workbook.xml");
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const actualXmlPath = resolve(actualFileDir, "xl/workbook.xml");
    const actualXml = readFileSync(actualXmlPath, "utf8");

    const expectedObj = parser.parse(expectedXml);
    const actualObj = parser.parse(actualXml);

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

    const expectedObj = parser.parse(expectedRels);
    const actualObj = parser.parse(actualRels);

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

    const expectedObj = parser.parse(expectedXml);
    const actualObj = parser.parse(actualXml);

    // It should be a problem-free difference.
    deletePropertyFromObject(
      expectedObj,
      "worksheet.sheetViews.sheetView.selection"
    );
    deletePropertyFromObject(
      actualObj,
      "worksheet.sheetViews.sheetView.selection"
    );

    // It maybe be a problem-free difference.
    deletePropertyFromObject(expectedObj, "worksheet.cols");

    expect(actualObj).toEqual(expectedObj);
  });
});
