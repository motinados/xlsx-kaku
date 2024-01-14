import { Borders } from "../src/borders";

describe("Borders", () => {
  test("getBorderId", () => {
    const borders = new Borders();
    expect(borders.getBorderId({})).toBe(0);
    expect(
      borders.getBorderId({ left: { style: "thin", color: "FF0000" } })
    ).toBe(1);
    expect(
      borders.getBorderId({ left: { style: "thin", color: "FFFFFF" } })
    ).toBe(2);
    expect(
      borders.getBorderId({ left: { style: "thin", color: "FF0000" } })
    ).toBe(1);
  });

  test("makeXml", () => {
    const borders = new Borders();
    expect(borders.makeXml()).toBe(
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    );

    borders.getBorderId({ left: { style: "thin", color: "FF000000" } });
    expect(borders.makeXml()).toBe(
      '<borders count="2">' +
        "<border><left/><right/><top/><bottom/><diagonal/></border>" +
        '<border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/></border>' +
        "</borders>"
    );

    borders.getBorderId({ right: { style: "thin", color: "FF000000" } });
    expect(borders.makeXml()).toBe(
      '<borders count="3">' +
        "<border><left/><right/><top/><bottom/><diagonal/></border>" +
        '<border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/></border>' +
        '<border><left/><right style="thin"><color rgb="FF000000"/></right><top/><bottom/><diagonal/></border>' +
        "</borders>"
    );

    borders.getBorderId({ top: { style: "thin", color: "FF000000" } });
    expect(borders.makeXml()).toBe(
      '<borders count="4">' +
        "<border><left/><right/><top/><bottom/><diagonal/></border>" +
        '<border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/></border>' +
        '<border><left/><right style="thin"><color rgb="FF000000"/></right><top/><bottom/><diagonal/></border>' +
        '<border><left/><right/><top style="thin"><color rgb="FF000000"/></top><bottom/><diagonal/></border>' +
        "</borders>"
    );

    borders.getBorderId({ bottom: { style: "thin", color: "FF000000" } });
    expect(borders.makeXml()).toBe(
      '<borders count="5">' +
        "<border><left/><right/><top/><bottom/><diagonal/></border>" +
        '<border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/></border>' +
        '<border><left/><right style="thin"><color rgb="FF000000"/></right><top/><bottom/><diagonal/></border>' +
        '<border><left/><right/><top style="thin"><color rgb="FF000000"/></top><bottom/><diagonal/></border>' +
        '<border><left/><right/><top/><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>' +
        "</borders>"
    );
  });
});
