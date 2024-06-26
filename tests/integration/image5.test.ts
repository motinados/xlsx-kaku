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

describe("inserting duplicated images", () => {
  const testName = "image5";

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

    const image1 = new Uint8Array(readFileSync("tests/assets/ufo_ushi.png"));
    const image2 = new Uint8Array(
      readFileSync("tests/assets/ufo_mikakunin_hikou_buttai.png")
    );

    await ws.insertImage({
      displayName: "ufo_ushi",
      extension: "png",
      data: image1,
      from: {
        col: 0,
        row: 0,
      },
      width: 180,
      height: 180,
    });

    await ws.insertImage({
      displayName: "ufo_mikakunin_hikou_buttai",
      extension: "png",
      data: image2,
      from: {
        col: 0,
        row: 14,
      },
      width: 180,
      height: 180,
    });

    await ws.insertImage({
      displayName: "ufo_ushi",
      extension: "png",
      data: image1,
      from: {
        col: 0,
        row: 29,
      },
      width: 180,
      height: 180,
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
    // It should be a problem-free difference.
    deletePropertyFromObject(expectedObj, "styleSheet.dxfs");
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

    // It may be a problem-free difference.
    deletePropertyFromObject(expectedObj, "workbook.workbookPr");
    deletePropertyFromObject(actualObj, "workbook.workbookPr");

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

  test("drawing1.xml.rels", () => {
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
      "xl/drawings/_rels/drawing1.xml.rels"
    );
    const expectedRels = readFileSync(expectedRelsPath, "utf8");
    const actualRelsPath = resolve(
      actualFileDir,
      "xl/drawings/_rels/drawing1.xml.rels"
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

  test("drawing1.xml", () => {
    const expectedXmlPath = resolve(
      expectedFileDir,
      "xl/drawings/drawing1.xml"
    );
    const expectedXml = readFileSync(expectedXmlPath, "utf8");
    const actualXmlPath = resolve(actualFileDir, "xl/drawings/drawing1.xml");
    const actualXml = readFileSync(actualXmlPath, "utf8");

    const expectedObj = parseXml(expectedXml);
    const actualObj = parseXml(actualXml);

    // expectedObj is twoCellAnchor, but replace it with oneCellAnchor for testing.
    const body = expectedObj["xdr:wsDr"]["xdr:twoCellAnchor"];
    expectedObj["xdr:wsDr"]["xdr:oneCellAnchor"] = body;

    for (let i = 0; i < body.length; i++) {
      const ext = body[i]["xdr:pic"]["xdr:spPr"]["a:xfrm"]["a:ext"];
      body[i]["xdr:ext"] = ext;

      // Not required for oneCellAnchor
      deletePropertyFromObject(body[i], "xdr:to");
    }

    deletePropertyFromObject(expectedObj, "xdr:wsDr.xdr:twoCellAnchor");

    // It should be a problem-free difference.
    for (const obj of expectedObj["xdr:wsDr"]["xdr:oneCellAnchor"]) {
      deletePropertyFromObject(obj, "xdr:pic.xdr:nvPicPr.xdr:cNvPr.@_name");
    }

    for (const obj of actualObj["xdr:wsDr"]["xdr:oneCellAnchor"]) {
      deletePropertyFromObject(obj, "xdr:pic.xdr:nvPicPr.xdr:cNvPr.@_name");
    }

    // It should be a problem-free difference.
    for (const obj of expectedObj["xdr:wsDr"]["xdr:oneCellAnchor"]) {
      const ext =
        obj["xdr:pic"]["xdr:nvPicPr"]["xdr:cNvPr"]["a:extLst"]["a:ext"];

      if (Array.isArray(ext)) {
        for (const e of ext) {
          if (e["a16:creationId"]) {
            deletePropertyFromObject(e["a16:creationId"], "@_id");
          } else if (e["a16:predDERef"]) {
            deletePropertyFromObject(e["a16:predDERef"], "@_pred");
          }
        }
      } else {
        deletePropertyFromObject(ext, "a16:creationId.@_id");
      }
    }

    // It should be a problem-free difference.
    for (const obj of actualObj["xdr:wsDr"]["xdr:oneCellAnchor"]) {
      const ext =
        obj["xdr:pic"]["xdr:nvPicPr"]["xdr:cNvPr"]["a:extLst"]["a:ext"];

      if (Array.isArray(ext)) {
        for (const e of ext) {
          if (e["a16:creationId"]) {
            deletePropertyFromObject(e["a16:creationId"], "@_id");
          } else if (e["a16:predDERef"]) {
            deletePropertyFromObject(e["a16:predDERef"], "@_pred");
          }
        }
      } else {
        deletePropertyFromObject(ext, "a16:creationId.@_id");
      }
    }

    // It may be a problem-free difference.
    const expectedImage2Obj = expectedObj["xdr:wsDr"]["xdr:oneCellAnchor"][1];
    deletePropertyFromObject(
      expectedImage2Obj["xdr:pic"]["xdr:spPr"]["a:xfrm"]["a:off"],
      "@_y"
    );
    const actualImage2Obj = actualObj["xdr:wsDr"]["xdr:oneCellAnchor"][1];
    deletePropertyFromObject(
      actualImage2Obj["xdr:pic"]["xdr:spPr"]["a:xfrm"]["a:off"],
      "@_y"
    );

    // It may be a problem-free difference.
    const expectedImage3Obj = expectedObj["xdr:wsDr"]["xdr:oneCellAnchor"][2];
    deletePropertyFromObject(
      expectedImage3Obj["xdr:pic"]["xdr:spPr"]["a:xfrm"]["a:off"],
      "@_y"
    );
    const actualImage3Obj = actualObj["xdr:wsDr"]["xdr:oneCellAnchor"][2];
    deletePropertyFromObject(
      actualImage3Obj["xdr:pic"]["xdr:spPr"]["a:xfrm"]["a:off"],
      "@_y"
    );

    expect(actualObj).toEqual(expectedObj);
  });

  test("media", () => {
    const expectedFiles = listFiles(resolve(expectedFileDir, "xl/media"));
    const actualFiles = listFiles(resolve(actualFileDir, "xl/media"));

    const expectedSubPaths = expectedFiles.map((it) =>
      removeBasePath(it, expectedFileDir)
    );
    const actualSubPaths = actualFiles.map((it) =>
      removeBasePath(it, actualFileDir)
    );

    expect(actualSubPaths).toEqual(expectedSubPaths);
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

    expect(actualObj).toEqual(expectedObj);
  });
});
