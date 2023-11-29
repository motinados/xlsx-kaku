import { stringifySorted } from "./fonts";

type CellXf = {
  fillId: number;
  fontId: number;
  borderId: number;
  numFmtId: number;
};

export class CellXfs {
  private cellXfs = new Map<string, number>([
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

  getCellXfId(cellXf: CellXf): number {
    const key = stringifySorted(cellXf);
    const id = this.cellXfs.get(key);
    if (id !== undefined) {
      return id;
    }

    const cellXfId = this.cellXfs.size;
    this.cellXfs.set(key, cellXfId);
    return cellXfId;
  }
}
