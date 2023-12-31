import { basename, extname, resolve } from "node:path";
import { listFiles, removeBasePath, unzip } from "../helper/helper";
import { Workbook } from "../../src";

const XLSX_Dir = "tests/xlsx";
const OUTPUT_DIR = "tests/temp/mergeCells/output";
const EXPECTED_UNZIPPED_DIR = "tests/temp/mergeCells/expected";
const ACTUAL_UNZIPPED_DIR = "tests/temp/mergeCells/actuall";

describe("string", () => {
  let xlsxBaseName: string;
  let expectedFileDir: string;
  let actualFileDir: string;
  let outputPath: string;

  beforeAll(async () => {
    const filepath = resolve(XLSX_Dir, "mergeCells.xlsx");

    const extension = extname(filepath);
    xlsxBaseName = basename(filepath, extension);
    expectedFileDir = resolve(EXPECTED_UNZIPPED_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 1 });
    ws.setCell(1, 0, { type: "number", value: 2 });
    ws.setMergeCell({ ref: "A1:C1" });
    ws.setMergeCell({ ref: "A2:A4" });

    outputPath = resolve(OUTPUT_DIR, "mergeCells.xlsx");
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
