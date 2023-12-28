import { NumberFormats } from "../src/numberFormats";
describe("Styles", () => {
  test("default getNumFmtId", () => {
    const numberFormats = new NumberFormats();
    expect(numberFormats.getNumFmtId("0")).toBe(1);
    expect(numberFormats.getNumFmtId("0.00")).toBe(2);
    expect(numberFormats.getNumFmtId("#,##0")).toBe(3);
    expect(numberFormats.getNumFmtId("#,##0.00")).toBe(4);
    expect(numberFormats.getNumFmtId("0%")).toBe(9);
    expect(numberFormats.getNumFmtId("0.00%")).toBe(10);
    expect(numberFormats.getNumFmtId("0.00E+00")).toBe(11);
    expect(numberFormats.getNumFmtId("# ?/?")).toBe(12);
    expect(numberFormats.getNumFmtId("# ??/??")).toBe(13);
    expect(numberFormats.getNumFmtId("mm-dd-yy")).toBe(14);
    expect(numberFormats.getNumFmtId("d-mmm-yy")).toBe(15);
    expect(numberFormats.getNumFmtId("d-mmm")).toBe(16);
    expect(numberFormats.getNumFmtId("mmm-yy")).toBe(17);
    expect(numberFormats.getNumFmtId("h:mm AM/PM")).toBe(18);
    expect(numberFormats.getNumFmtId("h:mm:ss AM/PM")).toBe(19);
    expect(numberFormats.getNumFmtId("h:mm")).toBe(20);
    expect(numberFormats.getNumFmtId("h:mm:ss")).toBe(21);
    expect(numberFormats.getNumFmtId("m/d/yy h:mm")).toBe(22);
    expect(numberFormats.getNumFmtId("#,##0 ;(#,##0)")).toBe(37);
    expect(numberFormats.getNumFmtId("#,##0 ;[Red](#,##0)")).toBe(38);
    expect(numberFormats.getNumFmtId("#,##0.00;(#,##0.00)")).toBe(39);
    expect(numberFormats.getNumFmtId("#,##0.00;[Red](#,##0.00)")).toBe(40);
    expect(numberFormats.getNumFmtId("mm:ss")).toBe(45);
    expect(numberFormats.getNumFmtId("[h]:mm:ss")).toBe(46);
    expect(numberFormats.getNumFmtId("mmss.0")).toBe(47);
    expect(numberFormats.getNumFmtId("##0.0E+0")).toBe(48);
    expect(numberFormats.getNumFmtId("@")).toBe(49);
  });

  test("custom getNumFmtId", () => {
    const numberFormats = new NumberFormats();
    expect(numberFormats.getNumFmtId("yyyy-mm-dd")).toBe(176);
    expect(numberFormats.getNumFmtId("mm-yyyy-dd")).toBe(177);
  });

  test("makeXml", () => {
    const numberFormats = new NumberFormats();
    numberFormats.getNumFmtId("yyyy-mm-dd");
    numberFormats.getNumFmtId("yyyy/m/d h:mm");
    const xml = numberFormats.makeXml();
    expect(xml).toBe(
      `<numFmts count="2"><numFmt numFmtId="176" formatCode="yyyy\\-mm\\-dd;@"/><numFmt numFmtId="177" formatCode="yyyy/m/d\\ h:mm;@"/></numFmts>`
    );
  });
});
