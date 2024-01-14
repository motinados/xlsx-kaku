import { Fonts } from "../src/fonts";

describe("Fonts", () => {
  test("getFontId", () => {
    const fonts = new Fonts();
    expect(
      fonts.getFontId({
        name: "Calibri",
        color: "000000",
        size: 11,
      })
    ).toBe(0);

    expect(
      fonts.getFontId({
        size: 11,
        name: "Calibri",
        color: "000000",
      })
    ).toBe(0);

    expect(
      fonts.getFontId({
        name: "Arial",
        size: 12,
        color: "FF0000",
      })
    ).toBe(1);
  });

  test("makeXml", () => {
    const fonts = new Fonts();
    fonts.getFontId({
      name: "Calibri",
      color: "000000",
      size: 11,
    });
    fonts.getFontId({
      name: "Arial",
      color: "FF0000",
      size: 12,
    });
    expect(fonts.makeXml()).toBe(
      '<fonts count="2"><font><sz val="11"/><color rgb="000000"/><name val="Calibri"/></font><font><sz val="12"/><color rgb="FF0000"/><name val="Arial"/></font></fonts>'
    );
  });

  test("makeXml with bold", () => {
    const fonts = new Fonts();
    fonts.getFontId({
      name: "Calibri",
      color: "000000",
      size: 11,
      bold: true,
    });
    expect(fonts.makeXml()).toBe(
      '<fonts count="2">' +
        '<font><sz val="11"/><color rgb="000000"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="000000"/><name val="Calibri"/></font>' +
        "</fonts>"
    );
  });

  test("makeXml with underline", () => {
    const fonts = new Fonts();
    fonts.getFontId({
      name: "Calibri",
      color: "000000",
      size: 11,
      underline: "single",
    });
    fonts.getFontId({
      name: "Calibri",
      color: "000000",
      size: 11,
      underline: "double",
    });
    expect(fonts.makeXml()).toBe(
      '<fonts count="3">' +
        '<font><sz val="11"/><color rgb="000000"/><name val="Calibri"/></font>' +
        '<font><u/><sz val="11"/><color rgb="000000"/><name val="Calibri"/></font>' +
        '<font><u val="double"/><sz val="11"/><color rgb="000000"/><name val="Calibri"/></font>' +
        "</fonts>"
    );
  });
});
