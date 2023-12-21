import path, { basename, extname } from "node:path";
import { listFiles, removeBasePath, unzip } from "./helper/helper";
import { Workbook } from "../src/index";
import { rmSync } from "node:fs";

const OUTPUT_DIR = "tests/output";
const EXPECTED_FILE_DIR = "tests/expected";
const ACTUAL_FILE_DIR = "tests/actuall";

describe("number", () => {
  test("number", async () => {
    const filepath = path.resolve("tests/xlsx/number.xlsx");

    const extension = extname(filepath);
    const xlsxBaseName = basename(filepath, extension);
    const expectedFileDir = path.resolve(EXPECTED_FILE_DIR, xlsxBaseName);
    await unzip(filepath, expectedFileDir);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "number", value: 15 });
    const outputPath = path.resolve(OUTPUT_DIR, "number.xlsx");
    await wb.save(outputPath);

    const actualFileDir = path.resolve(ACTUAL_FILE_DIR, xlsxBaseName);
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

    rmSync(OUTPUT_DIR, { recursive: true });
    rmSync(EXPECTED_FILE_DIR, { recursive: true });
    rmSync(ACTUAL_FILE_DIR, { recursive: true });
  });
});
