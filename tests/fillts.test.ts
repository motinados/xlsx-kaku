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
});
