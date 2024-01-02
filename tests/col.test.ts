import { Col, combineColProps } from "../src/col";

describe("col", () => {
  test("combineColProps", () => {
    const cols: Col[] = [
      { min: 1, max: 2, width: 10 },
      { min: 1, max: 2, style: { alignment: { horizontal: "center" } } },
      { min: 3, max: 4, width: 20 },
      { min: 3, max: 4, style: { alignment: { vertical: "top" } } },
    ];
    const combinedCols = combineColProps(cols);
    expect(combinedCols).toEqual([
      {
        min: 1,
        max: 2,
        width: 10,
        style: { alignment: { horizontal: "center" } },
      },
      { min: 3, max: 4, width: 20, style: { alignment: { vertical: "top" } } },
    ]);
  });
});
