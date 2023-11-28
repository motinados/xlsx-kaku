import { getFontId } from "../src/fonts";

describe("Fonts", () => {
  test("getFontId", () => {
    expect(
      getFontId({
        name: "游ゴシック",
        color: "FF0000",
        size: 11,
      })
    ).toBe(0);

    expect(
      getFontId({
        size: 11,
        name: "游ゴシック",
        color: "FF0000",
      })
    ).toBe(0);

    expect(
      getFontId({
        name: "Arial",
        size: 12,
        color: "FF0000",
      })
    ).toBe(1);
  });
});
