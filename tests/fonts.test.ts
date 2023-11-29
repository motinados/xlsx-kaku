import { Fonts } from "../src/fonts";

describe("Fonts", () => {
  test("getFontId", () => {
    const fonts = new Fonts();
    expect(
      fonts.getFontId({
        name: "游ゴシック",
        color: "FF0000",
        size: 11,
      })
    ).toBe(0);

    expect(
      fonts.getFontId({
        size: 11,
        name: "游ゴシック",
        color: "FF0000",
      })
    ).toBe(0);

    expect(
      fonts.getFontId({
        name: "Arial",
        size: 12,
        color: "FF0000",
      })
    ).toBe(1);
  });
});
