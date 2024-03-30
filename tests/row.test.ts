import { Worksheet } from "../src";

describe("row", () => {
  test("setRowOpts", () => {
    const ws = new Worksheet("Sheet1");

    ws.setRowOpts({ index: 0, height: 10 });
    expect(ws.rowOptsMap).toStrictEqual(
      new Map([
        [
          0,
          {
            index: 0,
            height: 10,
          },
        ],
      ])
    );

    ws.setRowOpts({ index: 1, height: 20 });
    expect(ws.rowOptsMap).toStrictEqual(
      new Map([
        [
          0,
          {
            index: 0,
            height: 10,
          },
        ],
        [
          1,
          {
            index: 1,
            height: 20,
          },
        ],
      ])
    );

    // setRowOpts overwrite existing props
    ws.setRowOpts({
      index: 0,
      style: { alignment: { horizontal: "center" } },
    });
    expect(ws.rowOptsMap).toStrictEqual(
      new Map([
        [
          0,
          {
            index: 0,
            style: { alignment: { horizontal: "center" } },
          },
        ],
        [
          1,
          {
            index: 1,
            height: 20,
          },
        ],
      ])
    );

    ws.setRowOpts({ index: 1, style: { alignment: { vertical: "top" } } });
    expect(ws.rowOptsMap).toStrictEqual(
      new Map([
        [
          0,
          {
            index: 0,
            style: { alignment: { horizontal: "center" } },
          },
        ],
        [
          1,
          {
            index: 1,
            style: { alignment: { vertical: "top" } },
          },
        ],
      ])
    );
  });
});
