import { MergeCells } from "../src/mergeCells";
describe("mergeCell", () => {
  test("should be able to create a mergeCell", () => {
    const mergeCells = new MergeCells();
    expect(mergeCells).toBeInstanceOf(MergeCells);
  });

  test("should be able to add a mergeCell", () => {
    const mergeCells = new MergeCells();
    expect(mergeCells.count).toBe(0);
    mergeCells.addMergeCell({ ref: "A1:B2" });
    expect(mergeCells.count).toBe(1);
    mergeCells.addMergeCell({ ref: "C1:D3" });
    expect(mergeCells.count).toBe(2);
    expect(mergeCells.mergeCells).toStrictEqual([
      { ref: "A1:B2" },
      { ref: "C1:D3" },
    ]);
  });
});
