import { Hyperlinks } from "../src/hyperlinks";

describe("hyperlinks", () => {
  test("hyperlink with a linkType of 'external'", () => {
    const hyperlinks = new Hyperlinks();
    hyperlinks.addHyperlink({
      linkType: "external",
      ref: "A1",
      rid: "rId1",
      uuid: "00000000-0000-0000-0000-000000000000",
    });

    expect(hyperlinks.getHyperlinks()).toEqual([
      {
        linkType: "external",
        ref: "A1",
        rid: "rId1",
        uuid: "00000000-0000-0000-0000-000000000000",
      },
    ]);

    expect(hyperlinks.makeXML()).toEqual(
      '<hyperlinks><hyperlink ref="A1" r:id="rId1" xr:uid="{00000000-0000-0000-0000-000000000000}"/></hyperlinks>'
    );
  });

  test("hyperlink with a linkType of 'internal'", () => {
    const hyperlinks = new Hyperlinks();
    hyperlinks.addHyperlink({
      linkType: "internal",
      ref: "A1",
      location: "Sheet1!A1",
      display: "toA1",
      uuid: "00000000-0000-0000-0000-000000000000",
    });
    hyperlinks.addHyperlink({
      linkType: "internal",
      ref: "A2",
      location: "A2",
      display: "toA2",
      uuid: "00000000-0000-0000-0000-000000000001",
    });

    expect(hyperlinks.getHyperlinks()).toEqual([
      {
        linkType: "internal",
        ref: "A1",
        location: "Sheet1!A1",
        display: "toA1",
        uuid: "00000000-0000-0000-0000-000000000000",
      },
      {
        linkType: "internal",
        ref: "A2",
        location: "A2",
        display: "toA2",
        uuid: "00000000-0000-0000-0000-000000000001",
      },
    ]);

    expect(hyperlinks.makeXML()).toEqual(
      "<hyperlinks>" +
        `<hyperlink ref="A1" location="'Sheet1'!A1" display="toA1" xr:uid="{00000000-0000-0000-0000-000000000000}"/>` +
        `<hyperlink ref="A2" location="A2" display="toA2" xr:uid="{00000000-0000-0000-0000-000000000001}"/>` +
        "</hyperlinks>"
    );
  });

  test("hyperlink with a linkType of 'email'", () => {
    const hyperlinks = new Hyperlinks();
    hyperlinks.addHyperlink({
      linkType: "email",
      ref: "A1",
      rid: "rId1",
      uuid: "00000000-0000-0000-0000-000000000000",
    });
    hyperlinks.addHyperlink({
      linkType: "email",
      ref: "A2",
      rid: "rId2",
      uuid: "00000000-0000-0000-0000-000000000001",
    });

    expect(hyperlinks.getHyperlinks()).toEqual([
      {
        linkType: "email",
        ref: "A1",
        rid: "rId1",
        uuid: "00000000-0000-0000-0000-000000000000",
      },
      {
        linkType: "email",
        ref: "A2",
        rid: "rId2",
        uuid: "00000000-0000-0000-0000-000000000001",
      },
    ]);

    expect(hyperlinks.makeXML()).toEqual(
      '<hyperlinks><hyperlink ref="A1" r:id="rId1" xr:uid="{00000000-0000-0000-0000-000000000000}"/><hyperlink ref="A2" r:id="rId2" xr:uid="{00000000-0000-0000-0000-000000000001}"/></hyperlinks>'
    );
  });

  test("hyperlinks", () => {
    const hyperlinks = new Hyperlinks();
    hyperlinks.addHyperlink({
      linkType: "external",
      ref: "A1",
      rid: "rId1",
      uuid: "00000000-0000-0000-0000-000000000000",
    });
    hyperlinks.addHyperlink({
      linkType: "external",
      ref: "A2",
      rid: "rId2",
      uuid: "00000000-0000-0000-0000-000000000001",
    });

    expect(hyperlinks.getHyperlinks()).toEqual([
      {
        linkType: "external",
        ref: "A1",
        rid: "rId1",
        uuid: "00000000-0000-0000-0000-000000000000",
      },
      {
        linkType: "external",
        ref: "A2",
        rid: "rId2",
        uuid: "00000000-0000-0000-0000-000000000001",
      },
    ]);

    expect(hyperlinks.makeXML()).toEqual(
      '<hyperlinks><hyperlink ref="A1" r:id="rId1" xr:uid="{00000000-0000-0000-0000-000000000000}"/><hyperlink ref="A2" r:id="rId2" xr:uid="{00000000-0000-0000-0000-000000000001}"/></hyperlinks>'
    );
  });
});
