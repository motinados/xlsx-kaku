// import { rmSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import { XMLParser } from "fast-xml-parser";
import { listFiles, removeBasePath, unzip } from "../helper/helper";
import { Workbook } from "../../src";
import { readFileSync } from "node:fs";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/boolean/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/boolean/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/boolean/actuall";

const parser = new XMLParser({ ignoreAttributes: false });

describe("string", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = resolve(XLSX_Dir, "boolean.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "boolean", value: true });
    ws.setCell(0, 1, { type: "boolean", value: false });
    ws.setCell(1, 0, { type: "boolean", value: false });
    ws.setCell(1, 1, { type: "boolean", value: true });
    outputPath = resolve(OUTPUT_DIR, "boolean.xlsx");
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

    const expectedObj = parser.parse(expected);
    const actualObj = parser.parse(actual);

    expect(actualObj).toEqual(expectedObj);
  });
});
