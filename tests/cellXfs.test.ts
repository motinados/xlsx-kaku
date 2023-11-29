import { CellXfs } from "../src/cellXfs";

describe("Styles", () => {
  test("getStyleId", () => {
    const styles = new CellXfs();
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(0);
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 1,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(1);
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 0,
        borderId: 1,
        numFmtId: 0,
      })
    ).toBe(2);
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(0);
  });
});
