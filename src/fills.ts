import { stringifySorted } from "./fonts";

type Fill =
  | {
      patternType: "none";
    }
  | {
      patternType: "gray125";
    }
  | {
      patternType: "solid";
      fgColor: string;
    };

export class Fills {
  private fills = new Map<string, number>([
    [
      stringifySorted({
        patternType: "none",
      }),
      0,
    ],
    [
      stringifySorted({
        patternType: "gray125",
      }),
      1,
    ],
  ]);

  getFillId(fill: Fill): number {
    const key = stringifySorted(fill);
    const id = this.fills.get(key);
    if (id !== undefined) {
      return id;
    }

    const fillId = this.fills.size;
    this.fills.set(key, fillId);
    return fillId;
  }
}
