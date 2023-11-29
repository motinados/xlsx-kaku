import { Styles } from "../src/styles";

describe("Styles", () => {
  test("getStyleId", () => {
    const styles = new Styles();
    expect(
      styles.getStyleId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(0);
    expect(
      styles.getStyleId({
        fillId: 0,
        fontId: 1,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(1);
    expect(
      styles.getStyleId({
        fillId: 0,
        fontId: 0,
        borderId: 1,
        numFmtId: 0,
      })
    ).toBe(2);
    expect(
      styles.getStyleId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      })
    ).toBe(0);
  });
});
