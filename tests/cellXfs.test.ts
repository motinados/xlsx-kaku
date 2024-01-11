import { CellXfs } from "../src/cellXfs";

describe("Styles", () => {
  test("getCellXfs", () => {
    const styles = new CellXfs();
    expect(styles.count).toBe(1);
    expect(styles.cellXfs.size).toBe(1);
  });

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

  test("getStyleId with alignment", () => {
    const styles = new CellXfs();
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
        alignment: {
          horizontal: "center",
          vertical: "center",
        },
      })
    ).toBe(1);
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
        alignment: {
          horizontal: "center",
          vertical: "top",
        },
      })
    ).toBe(2);
    expect(
      styles.getCellXfId({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
        alignment: {
          horizontal: "center",
          vertical: "center",
        },
      })
    ).toBe(1);
  });

  test("makeXml", () => {
    const styles = new CellXfs();
    expect(styles.makeXml()).toBe(
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`
    );

    styles.getCellXfId({
      fillId: 0,
      fontId: 0,
      borderId: 0,
      numFmtId: 0,
    });

    expect(styles.makeXml()).toBe(
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`
    );

    styles.getCellXfId({
      fillId: 0,
      fontId: 0,
      borderId: 0,
      numFmtId: 1,
    });

    let expected = "";
    expected += '<cellXfs count="2">';
    expected +=
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    expected +=
      '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>';
    expected += "</cellXfs>";
    expect(styles.makeXml()).toBe(expected);
  });

  test("makeXml with alignment", () => {
    const styles = new CellXfs();
    expect(styles.makeXml()).toBe(
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`
    );

    styles.getCellXfId({
      fillId: 0,
      fontId: 0,
      borderId: 0,
      numFmtId: 0,
      alignment: {
        horizontal: "center",
        vertical: "center",
        textRotation: 135,
      },
    });

    let expected = "";
    expected += '<cellXfs count="2">';
    expected +=
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    expected +=
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">';
    expected +=
      '<alignment horizontal="center" vertical="center" textRotation="135"/>';
    expected += "</xf>";
    expected += "</cellXfs>";
    expect(styles.makeXml()).toBe(expected);
  });
});
