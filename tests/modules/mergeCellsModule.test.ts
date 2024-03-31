import { Worksheet } from "../../src";
import { mergeCellsModule } from "../../src/modules/mergeCellsModule";

describe("mergeCellsModule", () => {
  test("getMergeCells", () => {
    const module = mergeCellsModule();
    expect(module.getMergeCells()).toStrictEqual([]);

    const worksheet = new Worksheet("test");
    module.add(worksheet, { ref: "A1:B2" });
    module.add(worksheet, { ref: "C3:D4" });
    module.add(worksheet, { ref: "E5:F6" });

    expect(module.getMergeCells()).toStrictEqual([
      { ref: "A1:B2" },
      { ref: "C3:D4" },
      { ref: "E5:F6" },
    ]);

    expect(module.makeXmlElm()).toBe(
      `<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="C3:D4"/><mergeCell ref="E5:F6"/></mergeCells>`
    );
  });
});
