import {
  convColNameToColIndex,
  convColIndexToColName,
  createUuid,
  devideAddress,
  escapeXmlText,
  expandRange,
  isInRange,
  hasSheetName,
  getFirstAddress,
  isRange,
  quoteSheetName,
} from "../src/utils";

describe("utils", () => {
  test("convColumnToNumber", () => {
    expect(convColNameToColIndex("A")).toBe(0);
    expect(convColNameToColIndex("B")).toBe(1);
    expect(convColNameToColIndex("Z")).toBe(25);
    expect(convColNameToColIndex("AA")).toBe(26);
    expect(convColNameToColIndex("BC")).toBe(54);
  });

  test("convNumberToColumn", () => {
    expect(convColIndexToColName(0)).toBe("A");
    expect(convColIndexToColName(1)).toBe("B");
    expect(convColIndexToColName(25)).toBe("Z");
    expect(convColIndexToColName(26)).toBe("AA");
    expect(convColIndexToColName(54)).toBe("BC");
  });

  test("devideAddress", () => {
    expect(devideAddress("A1")).toStrictEqual(["A", 1]);
    expect(devideAddress("B2")).toStrictEqual(["B", 2]);
    expect(devideAddress("Z3")).toStrictEqual(["Z", 3]);
    expect(devideAddress("AA10")).toStrictEqual(["AA", 10]);
    expect(devideAddress("BCD99")).toStrictEqual(["BCD", 99]);
  });

  test("expandRange", () => {
    expect(expandRange("A1:A1")).toStrictEqual([[0, 0]]);
    expect(expandRange("A1:A2")).toStrictEqual([
      [0, 0],
      [0, 1],
    ]);
    expect(expandRange("A1:B2")).toStrictEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
    expect(expandRange("A1:C3")).toStrictEqual([
      [0, 0],
      [0, 1],
      [0, 2],
      [1, 0],
      [1, 1],
      [1, 2],
      [2, 0],
      [2, 1],
      [2, 2],
    ]);
  });

  test("isInRange", () => {
    expect(isInRange("A", 1, 1)).toBe(true);
    expect(isInRange("A", 1, 2)).toBe(true);
    expect(isInRange("A", 2, 2)).toBe(false);

    expect(isInRange("C", 1, 1)).toBe(false);
    expect(isInRange("C", 1, 2)).toBe(false);
    expect(isInRange("C", 3, 3)).toBe(true);
    expect(isInRange("C", 3, 4)).toBe(true);
  });

  test("hasSheetName", () => {
    expect(hasSheetName("Sheet1!A1")).toBe(true);
    expect(hasSheetName("A1")).toBe(false);
  });

  test("isRange", () => {
    expect(isRange("A1:A1")).toBe(true);
    expect(isRange("A1:A2")).toBe(true);
    expect(isRange("A1:B2")).toBe(true);
    expect(isRange("A1")).toBe(false);
  });

  test("getFirstAddress", () => {
    expect(getFirstAddress("A1:B2")).toBe("A1");
    expect(getFirstAddress("B2")).toBe("B2");
  });

  test("createUuid", () => {
    expect(createUuid()).toMatch(
      /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/
    );
  });

  test("escapeXmlText", () => {
    expect(escapeXmlText("Sheet1")).toBe("Sheet1");
    expect(escapeXmlText("a & b")).toBe("a &amp; b");
    expect(escapeXmlText("<tag>")).toBe("&lt;tag&gt;");
    // Quotes do not need to be escaped in a text node.
    expect(escapeXmlText(`"quoted"`)).toBe(`"quoted"`);
    expect(escapeXmlText("it's")).toBe("it's");
    // The ampersand must not be escaped twice.
    expect(escapeXmlText("&lt;")).toBe("&amp;lt;");
  });

  test("quoteSheetName", () => {
    expect(quoteSheetName("Sheet1")).toBe("Sheet1");
    expect(quoteSheetName("_sheet")).toBe("_sheet");
    expect(quoteSheetName("my.sheet")).toBe("my.sheet");

    expect(quoteSheetName("My Sheet")).toBe("'My Sheet'");
    expect(quoteSheetName("2024")).toBe("'2024'");
    expect(quoteSheetName("sales-2024")).toBe("'sales-2024'");
    expect(quoteSheetName("A&B")).toBe("'A&B'");

    // A name that would be ambiguous with a cell reference.
    expect(quoteSheetName("A1")).toBe("'A1'");
    expect(quoteSheetName("XFD1048576")).toBe("'XFD1048576'");

    // A single quote in the name is doubled.
    expect(quoteSheetName("Bob's")).toBe("'Bob''s'");
  });
});
