// import { readFileSync, rmSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import {
  //   deletePropertyFromObject,
  listFiles,
  //   parseXml,
  removeBasePath,
  unzip,
} from "../helper/helper";
import { Workbook } from "../../src";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/rows/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/rows/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/rows/actuall";

describe("string", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = resolve(XLSX_Dir, "rows.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 1 });
    ws.setCell(1, 0, { type: "number", value: 2 });
    ws.setCell(2, 0, { type: "number", value: 3 });
    ws.setCell(3, 0, { type: "number", value: 4 });
    ws.setCell(4, 0, { type: "number", value: 5 });
    ws.setCell(5, 0, { type: "number", value: 6 });

    ws.setRowHeight({ index: 1, height: 20.25 });
    ws.setRowHeight({ index: 2, height: 39.75 });
    ws.setRowHeight({ index: 3, height: 39.75 });
    ws.setRowHeight({ index: 4, height: 39.75 });

    outputPath = resolve(OUTPUT_DIR, "rows.xlsx");
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
});
