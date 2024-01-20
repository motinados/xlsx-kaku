import { Worksheet } from "../src";
import { genXlsx } from "../src/writer";

describe("writer", () => {
  test("genXlsx", () => {
    const ws = new Worksheet("Sheet1");
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = genXlsx([ws]);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });
});
