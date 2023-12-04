import { CellStyles } from "../src/cellStyles";
describe("cellStyle", () => {
  test("getCellStyleId", () => {
    const styles = new CellStyles();
    expect(
      styles.getCellStyleId({
        name: "標準",
        xfId: 0,
        builtinId: 0,
      })
    ).toBe(0);
    expect(
      styles.getCellStyleId({
        name: "Hyperlink",
        xfId: 1,
        uid: "{00000000-000B-0000-0000-000008000000}",
      })
    ).toBe(1);
  });
});
