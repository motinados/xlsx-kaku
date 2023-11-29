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
});
