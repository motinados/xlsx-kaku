import { Worksheet } from "../src/worksheet";
import { createExcelFiles } from "../src/writer";

describe("string", () => {
  test("string", () => {
    const ws = new Worksheet("Sheet1");
    ws.sheetData = [[{ type: "string", value: "Hello" }]];

    const result = createExcelFiles([ws]);

    expect(result.styleMappers.cellXfs.count).toBe(1);
    expect(result.styleMappers.cellXfs.cellXfs).toEqual(
      new Map([['{"borderId":0,"fillId":0,"fontId":0,"numFmtId":0}', 0]])
    );

    expect(result.styleMappers.sharedStrings.count).toBe(1);
    expect(result.styleMappers.sharedStrings.uniqueCount).toBe(1);
    expect(result.sharedStringsXml).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' +
        "<si><t>Hello</t></si>" +
        "</sst>"
    );
  });
});
