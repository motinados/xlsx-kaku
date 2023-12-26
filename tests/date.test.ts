// import { rmSync } from "node:fs";
import { readFileSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import { XMLParser } from "fast-xml-parser";
import { listFiles, removeBasePath, unzip } from "./helper/helper";
import { Workbook } from "../src";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/date/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/date/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/date/actuall";

const parser = new XMLParser({ ignoreAttributes: false });

describe("date", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = resolve(XLSX_Dir, "date.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, {
      type: "date",
      value: new Date("2023-12-27").toISOString(),
    });
    ws.setCell(0, 1, {
      type: "date",
      value: new Date("2023-12-28").toISOString(),
    });
    ws.setCell(1, 0, {
      type: "date",
      value: new Date("2023-12-26").toISOString(),
    });
    ws.setCell(1, 1, {
      type: "date",
      value: new Date("2023-12-25").toISOString(),
    });
    outputPath = resolve(OUTPUT_DIR, "date.xlsx");
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
