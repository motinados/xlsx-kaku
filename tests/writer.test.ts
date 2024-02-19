import { Worksheet } from "../src";
import { genXlsx, genXlsxSync } from "../src/writer";

describe("writer", () => {
  test("genXlsx", async () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = await genXlsx([ws]);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("genXlsxSync", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = genXlsxSync([ws]);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("A Error should occur when there is no sheet.", async () => {
    try {
      await genXlsx([]);
    } catch (e) {
      expect(e).toBeInstanceOf(Error);
    }
  });
});
