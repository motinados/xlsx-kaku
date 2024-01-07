import { Col, combineColProps } from "../src/col";

describe("col", () => {
  test("combineColProps", () => {
    const cols: Col[] = [
      { startIndex: 0, endIndex: 1, width: 10 },
      {
        startIndex: 0,
        endIndex: 1,
        style: { alignment: { horizontal: "center" } },
      },
      { startIndex: 2, endIndex: 3, width: 20 },
      { startIndex: 2, endIndex: 3, style: { alignment: { vertical: "top" } } },
    ];
    const combinedCols = combineColProps(cols);
    expect(combinedCols).toEqual([
      {
        startIndex: 0,
        endIndex: 1,
        width: 10,
        style: { alignment: { horizontal: "center" } },
      },
      {
        startIndex: 2,
        endIndex: 3,
        width: 20,
        style: { alignment: { vertical: "top" } },
      },
    ]);
  });
});
