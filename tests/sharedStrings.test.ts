import { SharedStrings } from "../src/sharedStrings";
import { makeSharedStringsXml } from "../src/xml/sharedStringsXml";
describe("SharedStrings", () => {
  test("should be able to create a sharedStrings", () => {
    const sharedStrings = new SharedStrings();
    expect(sharedStrings).toBeInstanceOf(SharedStrings);
  });

  test("should be able to get index", () => {
    const sharedStrings = new SharedStrings();
    expect(sharedStrings.getIndex("hello")).toBe(0);
    expect(sharedStrings.getIndex("world")).toBe(1);
    expect(sharedStrings.getIndex("hello")).toBe(0);
  });

  test("should be able to get values in order", () => {
    const sharedStrings = new SharedStrings();
    sharedStrings.getIndex("hello");
    sharedStrings.getIndex("world");
    sharedStrings.getIndex("hello");
    expect(sharedStrings.getValuesInOrder()).toEqual(["hello", "world"]);
  });

  test("makeSharedStringsXml escapes the characters reserved in xml", () => {
    const sharedStrings = new SharedStrings();
    sharedStrings.getIndex("R&D");
    sharedStrings.getIndex("<b>bold</b>");

    expect(makeSharedStringsXml(sharedStrings)).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
        "<si><t>R&amp;D</t></si>" +
        "<si><t>&lt;b&gt;bold&lt;/b&gt;</t></si>" +
        "</sst>"
    );
  });

  test("makeSharedStringsXml preserves the surrounding whitespace", () => {
    const sharedStrings = new SharedStrings();
    sharedStrings.getIndex(" hello ");
    sharedStrings.getIndex("world");

    expect(makeSharedStringsXml(sharedStrings)).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
        '<si><t xml:space="preserve"> hello </t></si>' +
        "<si><t>world</t></si>" +
        "</sst>"
    );
  });
});
