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

  test("makeXml", () => {
    const styles = new CellStyleXfs();
    expect(styles.makeXml()).toBe(
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    );

    styles.getCellStyleXfId({
      fillId: 0,
      fontId: 1,
      borderId: 0,
      numFmtId: 0,
    });
    expect(styles.makeXml()).toBe(
      '<cellStyleXfs count="2">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>' +
        "</cellStyleXfs>"
    );

    styles.getCellStyleXfId({
      fillId: 0,
      fontId: 0,
      borderId: 1,
      numFmtId: 0,
    });
    expect(styles.makeXml()).toBe(
      '<cellStyleXfs count="3">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>' +
        "</cellStyleXfs>"
    );

    styles.getCellStyleXfId({
      fillId: 0,
      fontId: 0,
      borderId: 0,
      numFmtId: 1,
    });
    expect(styles.makeXml()).toBe(
      '<cellStyleXfs count="4">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>' +
        '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>' +
        "</cellStyleXfs>"
    );

    styles.getCellStyleXfId({
      fillId: 1,
      fontId: 0,
      borderId: 0,
      numFmtId: 0,
    });
    expect(styles.makeXml()).toBe(
      '<cellStyleXfs count="5">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>' +
        '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" applyFill="1"/>' +
        "</cellStyleXfs>"
    );
  });
});
