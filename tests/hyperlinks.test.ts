import { Hyperlinks } from "../src/hyperlinks";

describe("hyperlinks", () => {
  test("hyperlink", () => {
    const hyperlinks = new Hyperlinks();
    hyperlinks.addHyperlink({
      ref: "A1",
      rid: "rId1",
      uuid: "00000000-0000-0000-0000-000000000000",
    });

    expect(hyperlinks.getHyperlinks()).toEqual([
      {
        ref: "A1",
        rid: "rId1",
        uuid: "00000000-0000-0000-0000-000000000000",
      },
    ]);

    expect(hyperlinks.makeXML()).toEqual(
      '<hyperlinks><hyperlink ref="A1" r:id="rId1" xr:uid="{00000000-0000-0000-0000-000000000000}"/></hyperlinks>'
    );
  });

  test("hyperlinks", () => {
    const hyperlinks = new Hyperlinks();
    hyperlinks.addHyperlink({
      ref: "A1",
      rid: "rId1",
      uuid: "00000000-0000-0000-0000-000000000000",
    });
    hyperlinks.addHyperlink({
      ref: "A2",
      rid: "rId2",
      uuid: "00000000-0000-0000-0000-000000000001",
    });

    expect(hyperlinks.getHyperlinks()).toEqual([
      {
        ref: "A1",
        rid: "rId1",
        uuid: "00000000-0000-0000-0000-000000000000",
      },
      {
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
