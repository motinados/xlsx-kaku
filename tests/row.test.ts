import { Worksheet } from "../src";

describe("row", () => {
  test("setRowProps", () => {
    const ws = new Worksheet("Sheet1");

    ws.setRowProps({ index: 0, height: 10 });
    expect(ws.rows).toStrictEqual(
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

    ws.setRowProps({ index: 1, height: 20 });
    expect(ws.rows).toStrictEqual(
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

    // setRowProps overwrite existing props
    ws.setRowProps({
      index: 0,
      style: { alignment: { horizontal: "center" } },
    });
    expect(ws.rows).toStrictEqual(
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

    ws.setRowProps({ index: 1, style: { alignment: { vertical: "top" } } });
    expect(ws.rows).toStrictEqual(
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
