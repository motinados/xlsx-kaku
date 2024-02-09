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

describe("conditional formatting for date", () => {
  const testName = "conditionalFormatting6";

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

    const dates = [
      "2024-02-05T00:00:00.000Z",
      "2024-02-06T00:00:00.000Z",
      "2024-02-07T00:00:00.000Z",
      "2024-02-08T00:00:00.000Z",
      "2024-02-09T00:00:00.000Z",
      "2024-02-10T00:00:00.000Z",
      "2024-02-11T00:00:00.000Z",
      "2024-02-12T00:00:00.000Z",
      "2024-02-13T00:00:00.000Z",
      "2024-02-14T00:00:00.000Z",
      "2024-02-15T00:00:00.000Z",
      "2024-02-16T00:00:00.000Z",
      "2024-02-17T00:00:00.000Z",
      "2024-02-18T00:00:00.000Z",
      "2024-02-19T00:00:00.000Z",
      "2024-02-20T00:00:00.000Z",
      "2024-02-21T00:00:00.000Z",
      "2024-02-22T00:00:00.000Z",
      "2024-02-23T00:00:00.000Z",
      "2024-02-24T00:00:00.000Z",
      "2024-02-25T00:00:00.000Z",
      "2024-02-26T00:00:00.000Z",
      "2024-02-27T00:00:00.000Z",
      "2024-02-28T00:00:00.000Z",
      "2024-02-29T00:00:00.000Z",
      "2024-03-01T00:00:00.000Z",
      "2024-03-02T00:00:00.000Z",
      "2024-03-03T00:00:00.000Z",
      "2024-03-04T00:00:00.000Z",
      "2024-03-05T00:00:00.000Z",
    ];

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 0, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "B1:B1048576",
      priority: 10,
      timePeriod: "today",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 1, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "A1:A1048576",
      priority: 9,
      timePeriod: "yesterday",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 2, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "C1:C1048576",
      priority: 8,
      timePeriod: "tomorrow",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 3, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "D1:D1048576",
      priority: 7,
      timePeriod: "last7Days",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 4, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "E1:E1048576",
      priority: 6,
      timePeriod: "lastWeek",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 5, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "F1:F1048576",
      priority: 5,
      timePeriod: "thisWeek",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 6, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "G1:G1048576",
      priority: 4,
      timePeriod: "nextWeek",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 7, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "H1:H1048576",
      priority: 3,
      timePeriod: "lastMonth",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 8, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "I1:I1048576",
      priority: 2,
      timePeriod: "thisMonth",
      style: {
        font: { color: "FF9C0006" },
        fill: { bgColor: "FFFFC7CE" },
      },
    });

    for (let i = 0; i < dates.length; i++) {
      ws.setCell(i, 9, {
        type: "date",
        value: dates[i]!,
        style: {
          numberFormat: { formatCode: "yyyy-mm-dd;@" },
        },
      });
    }

    ws.setConditionalFormatting({
      type: "timePeriod",
      sqref: "J1:J1048576",
      priority: 1,
      timePeriod: "nextMonth",
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
