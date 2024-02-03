import { Dxf } from "../src/dxf";

describe("dxf", () => {
  test("makeDxfXml", () => {
    const dxf = new Dxf();
    const xml = dxf.makeXml();
    expect(xml).toBe("");

    dxf.addStyle({
      font: { color: "FF9C0006" },
      fill: { bgColor: "FFFFC7CE" },
    });

    const expected = `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>`;
    const actual = dxf.makeXml();
    expect(actual).toBe(expected);
  });
});
