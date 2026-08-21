import {
  autoFilterModule,
  convRefToAbsolute,
  normalizeAutoFilterRef,
} from "../../src/modules/autoFilterModule";

describe("autoFilterModule", () => {
  test("getAutoFilter returns null until it is set", () => {
    const module = autoFilterModule();

    expect(module.getAutoFilter()).toBeNull();
    expect(module.makeXmlElm()).toBe("");
    expect(module.makeDefinedNameElm("Sheet1", 0)).toBe("");
  });

  test("set and makeXmlElm", () => {
    const module = autoFilterModule();
    module.set({ ref: "A1:C10" });

    expect(module.getAutoFilter()).toStrictEqual({ ref: "A1:C10" });
    expect(module.makeXmlElm()).toBe(`<autoFilter ref="A1:C10"/>`);
  });

  test("a worksheet has at most one auto filter, so set replaces it", () => {
    const module = autoFilterModule();
    module.set({ ref: "A1:C10" });
    module.set({ ref: "B2:D20" });

    expect(module.getAutoFilter()).toStrictEqual({ ref: "B2:D20" });
    expect(module.makeXmlElm()).toBe(`<autoFilter ref="B2:D20"/>`);
  });

  test("the ref is normalized when it is set", () => {
    const module = autoFilterModule();
    module.set({ ref: "C10:A1" });

    expect(module.getAutoFilter()).toStrictEqual({ ref: "A1:C10" });
  });

  test("makeDefinedNameElm", () => {
    const module = autoFilterModule();
    module.set({ ref: "A1:C10" });

    expect(module.makeDefinedNameElm("Sheet1", 0)).toBe(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$C$10</definedName>`
    );
    expect(module.makeDefinedNameElm("Sheet2", 1)).toBe(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Sheet2!$A$1:$C$10</definedName>`
    );
  });

  test("makeDefinedNameElm quotes a sheet name that is not an identifier", () => {
    const module = autoFilterModule();
    module.set({ ref: "A1:B2" });

    expect(module.makeDefinedNameElm("My Sheet", 0)).toBe(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'My Sheet'!$A$1:$B$2</definedName>`
    );
  });

  test("makeDefinedNameElm escapes a sheet name", () => {
    const module = autoFilterModule();
    module.set({ ref: "A1:B2" });

    expect(module.makeDefinedNameElm("A&B", 0)).toBe(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'A&amp;B'!$A$1:$B$2</definedName>`
    );
    expect(module.makeDefinedNameElm("a<b", 0)).toBe(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'a&lt;b'!$A$1:$B$2</definedName>`
    );
  });

  describe("normalizeAutoFilterRef", () => {
    test("keeps a range that already starts with the top-left address", () => {
      expect(normalizeAutoFilterRef("A1:C10")).toBe("A1:C10");
      expect(normalizeAutoFilterRef("B2:AA100")).toBe("B2:AA100");
    });

    test("reorders a range so that the top-left address comes first", () => {
      expect(normalizeAutoFilterRef("C10:A1")).toBe("A1:C10");
      expect(normalizeAutoFilterRef("A10:C1")).toBe("A1:C10");
      expect(normalizeAutoFilterRef("C1:A10")).toBe("A1:C10");
    });

    test("accepts the whole worksheet", () => {
      expect(normalizeAutoFilterRef("A1:XFD1048576")).toBe("A1:XFD1048576");
    });

    test.each([
      ["A1", "a single cell is not a range"],
      ["A:C", "a column range has no row numbers"],
      ["1:10", "a row range has no column names"],
      ["a1:c10", "lowercase column names"],
      ["A0:C10", "a row number starts from 1"],
      ["A1:C10:E20", "too many addresses"],
      ["", "an empty string"],
    ])("throws for %s (%s)", (ref) => {
      expect(() => normalizeAutoFilterRef(ref)).toThrow(
        "Invalid auto filter ref"
      );
    });

    test.each(["A1:XFE10", "A1:C1048577"])(
      "throws for %s because it is out of the worksheet",
      (ref) => {
        expect(() => normalizeAutoFilterRef(ref)).toThrow(
          "out of the range of a worksheet"
        );
      }
    );
  });

  test("convRefToAbsolute", () => {
    expect(convRefToAbsolute("A1:C10")).toBe("$A$1:$C$10");
    expect(convRefToAbsolute("AA100:AB200")).toBe("$AA$100:$AB$200");
  });
});
