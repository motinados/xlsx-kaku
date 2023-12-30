import { rmSync } from "node:fs";
import { basename, extname, resolve } from "node:path";
import { listFiles, removeBasePath, unzip } from "../helper/helper";
import { Workbook } from "../../src";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/cols/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/cols/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/cols/actuall";

describe("string", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = resolve(XLSX_Dir, "cols.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 1 });
    ws.setCell(0, 1, { type: "number", value: 2 });
    ws.setCell(0, 2, { type: "number", value: 3 });
    ws.setCell(0, 3, { type: "number", value: 4 });
    ws.setCell(0, 4, { type: "number", value: 5 });
    ws.setCell(0, 5, { type: "number", value: 6 });

    ws.setColWidth({ min: 2, max: 2, width: 25 });
    ws.setColWidth({ min: 3, max: 5, width: 6 });

    outputPath = resolve(OUTPUT_DIR, "cols.xlsx");
    await wb.save(outputPath);

    actualFileDir = resolve(ACTUAL_UNZIPPED_DIR, xlsxBaseName);
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
});
