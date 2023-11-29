import { stringifySorted } from "./fonts";

type Style = {
  fillId: number;
  fontId: number;
  borderId: number;
  numFmtId: number;
};

export class Styles {
  private styles = new Map<string, number>([
    [
      stringifySorted({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      }),
      0,
    ],
  ]);

  getStyleId(style: Style): number {
    const key = stringifySorted(style);
    const id = this.styles.get(key);
    if (id !== undefined) {
      return id;
    }

    const styleId = this.styles.size;
    this.styles.set(key, styleId);
    return styleId;
  }
}
