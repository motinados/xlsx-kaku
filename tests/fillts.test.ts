import { Fills } from "../src/fills";

describe("Fills", () => {
  test("getFillId", () => {
    const fills = new Fills();
    expect(fills.getFillId({ patternType: "none" })).toBe(0);
    expect(fills.getFillId({ patternType: "gray125" })).toBe(1);
    expect(fills.getFillId({ patternType: "solid", fgColor: "FF0000" })).toBe(
      2
    );
    expect(fills.getFillId({ patternType: "solid", fgColor: "FFFFFF" })).toBe(
      3
    );
  });

  test("makeXml", () => {
    const fills = new Fills();
    expect(fills.makeXml()).toBe(
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>`
    );

    fills.getFillId({ patternType: "solid", fgColor: "FF0000" });
    expect(fills.makeXml()).toBe(
      `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0000"/><bgColor indexed="64"/></patternFill></fill></fills>`
    );
  });
});
