import { CellStyleXfs } from "../src/cellStyleXfs";

describe("cellStyleXfs", () => {
  test("getStyleId", () => {
    const styles = new CellStyleXfs();
    expect(
      styles.getCellStyleXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(0);
    expect(
      styles.getCellStyleXfId({
        fillId: 0,
        fontId: 1,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(1);
    expect(
      styles.getCellStyleXfId({
        fillId: 0,
        fontId: 0,
        borderId: 1,
        numFmtId: 0,
      })
    ).toBe(2);
    expect(
      styles.getCellStyleXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(0);
  });
});
