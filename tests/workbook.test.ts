import fs from "node:fs";
import path from "node:path";
import { Workbook } from "../src/workbook";

describe("workbook", () => {
  test("save", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.sheetData = [[{ type: "string", value: "Hello" }]];
    const testdir = "testdir";
    const filepath = path.join(testdir, "a/b/test.xlsx");
    await wb.save(filepath);
    expect(fs.existsSync(filepath)).toBe(true);
    fs.rmSync(testdir, { recursive: true });
  });
});
