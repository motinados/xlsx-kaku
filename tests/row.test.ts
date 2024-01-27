import { Row, combineRowProps } from "../src/row";

describe("row", () => {
  test("combineRowProps", () => {
    const rows: Row[] = [
      { index: 0, height: 10 },
      { index: 0, style: { alignment: { horizontal: "center" } } },
      { index: 1, height: 20 },
      { index: 1, style: { alignment: { vertical: "top" } } },
      { index: 2, height: 30 },
    ];
    const combinedRows = combineRowProps(rows);
    expect(combinedRows).toStrictEqual(
      new Map([
        [
          0,
          {
            index: 0,
            height: 10,
            style: { alignment: { horizontal: "center" } },
          },
        ],
        [
          1,
          {
            index: 1,
            height: 20,
            style: { alignment: { vertical: "top" } },
          },
        ],
        [2, { index: 2, height: 30 }],
      ])
    );
  });
});
