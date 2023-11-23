type MergeCell = {
  ref: string;
};

export class MergeCells {
  private _mergeCells: MergeCell[] = [];
  constructor() {}

  get count() {
    return this._mergeCells.length;
  }

  get mergeCells() {
    return this._mergeCells;
  }

  addMergeCell(mergeCell: MergeCell) {
    // TODO: validate mergeCell
    this._mergeCells.push(mergeCell);
  }
}
