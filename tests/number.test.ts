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

describe("number", () => {
  test("number", async () => {
    const parser = new XMLParser({ ignoreAttributes: false });
    const builder = new XMLBuilder();
    const filepath = path.resolve(XLSX_Dir, "number.xlsx");

    const extension = extname(filepath);
    const xlsxBaseName = basename(filepath, extension);
    const expectedFileDir = path.resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 15 });
    const outputPath = path.resolve(OUTPUT_DIR, "number.xlsx");
    await wb.save(outputPath);

    const actualFileDir = path.resolve(ACTUAL_UNZIPPED_DIR, xlsxBaseName);
    await unzip(outputPath, actualFileDir);

    const expectedFiles = listFiles(expectedFileDir);
    const actualFiles = listFiles(actualFileDir);

    const expectedSubPaths = expectedFiles.map((it) =>
      removeBasePath(it, expectedFileDir)
    );
    const actualSubPaths = actualFiles.map((it) =>
      removeBasePath(it, actualFileDir)
    );

    expect(actualSubPaths).toEqual(expectedSubPaths);

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

    rmSync(OUTPUT_DIR, { recursive: true });
    rmSync(EXPECTED_UNZIPPED_DIR, { recursive: true });
    rmSync(ACTUAL_UNZIPPED_DIR, { recursive: true });
  });
});
