import { Worksheet } from "../src";
import { ImageStore } from "../src/imageStore";
import { genXlsx, genXlsxSync } from "../src/writer";

describe("writer", () => {
  test("genXlsx", async () => {
    const imageStore = new ImageStore();
    const ws = new Worksheet("Sheet1", imageStore);
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = await genXlsx([ws], imageStore);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("genXlsxSync", () => {
    const imageStore = new ImageStore();
    const ws = new Worksheet("Sheet1", imageStore);
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = genXlsxSync([ws], imageStore);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("A Error should occur when there is no sheet.", async () => {
    try {
      await genXlsx([], new ImageStore());
    } catch (e) {
      expect(e).toBeInstanceOf(Error);
    }
  });

  test("A Error should occur when there is no sheet.", () => {
    try {
      genXlsxSync([], new ImageStore());
    } catch (e) {
      expect(e).toBeInstanceOf(Error);
    }
  });
});
