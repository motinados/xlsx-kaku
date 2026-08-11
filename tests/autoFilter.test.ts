import { strFromU8, unzipSync } from "fflate";
import { Workbook, WorkbookS } from "../src";
import { parseXml } from "./helper/helper";

function unzipXlsx(xlsx: Uint8Array) {
  const files = unzipSync(xlsx);
  return {
    sheetXml(sheetNumber: number) {
      return strFromU8(files[`xl/worksheets/sheet${sheetNumber}.xml`]!);
    },
    workbookXml() {
      return strFromU8(files["xl/workbook.xml"]!);
    },
  };
}

describe("autoFilter", () => {
  test("setAutoFilter is reflected in the worksheet", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    expect(ws.autoFilter).toBeNull();

    ws.setAutoFilter({ ref: "A1:C10" });

    expect(ws.autoFilter).toStrictEqual({ ref: "A1:C10" });
  });

  test("setAutoFilter throws for an invalid ref", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    expect(() => ws.setAutoFilter({ ref: "A1" })).toThrow(
      "Invalid auto filter ref"
    );
    expect(ws.autoFilter).toBeNull();
  });

  test("the worksheet has an autoFilter element", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "string", value: "name" });
    ws.setCell(1, 0, { type: "string", value: "foo" });
    ws.setAutoFilter({ ref: "A1:A2" });

    const sheetXml = unzipXlsx(wb.generateXlsxSync()).sheetXml(1);
    const sheetObj = parseXml(sheetXml);

    expect(sheetObj.worksheet.autoFilter["@_ref"]).toBe("A1:A2");
  });

  test("the autoFilter element comes right after sheetData", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setCell(0, 0, { type: "string", value: "name" });
    ws.setMergeCell({ ref: "B1:C1" });
    ws.setAutoFilter({ ref: "A1:A2" });

    const sheetXml = unzipXlsx(wb.generateXlsxSync()).sheetXml(1);

    // In the schema of a worksheet, autoFilter comes after sheetData
    // and before mergeCells.
    expect(sheetXml).toContain(`</sheetData><autoFilter ref="A1:A2"/>`);
    expect(sheetXml.indexOf("<autoFilter")).toBeLessThan(
      sheetXml.indexOf("<mergeCells")
    );
  });

  test("the workbook has a _xlnm._FilterDatabase defined name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.setAutoFilter({ ref: "A1:C10" });

    const workbookXml = unzipXlsx(wb.generateXlsxSync()).workbookXml();
    const workbookObj = parseXml(workbookXml);

    const definedName = workbookObj.workbook.definedNames.definedName;
    expect(definedName["@_name"]).toBe("_xlnm._FilterDatabase");
    expect(definedName["@_localSheetId"]).toBe("0");
    expect(definedName["@_hidden"]).toBe("1");
    expect(definedName["#text"]).toBe("Sheet1!$A$1:$C$10");

    // definedNames comes after sheets and before extLst.
    expect(workbookXml.indexOf("<definedNames>")).toBeGreaterThan(
      workbookXml.indexOf("</sheets>")
    );
    expect(workbookXml.indexOf("<definedNames>")).toBeLessThan(
      workbookXml.indexOf("<extLst>")
    );
  });

  test("only the sheets with an auto filter get a defined name", () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");
    const ws3 = wb.addWorksheet("Sheet3");
    ws2.setAutoFilter({ ref: "A1:B5" });
    ws3.setAutoFilter({ ref: "C1:D5" });

    const unzipped = unzipXlsx(wb.generateXlsxSync());

    expect(unzipped.sheetXml(1)).not.toContain("<autoFilter");
    expect(unzipped.sheetXml(2)).toContain(`<autoFilter ref="A1:B5"/>`);
    expect(unzipped.sheetXml(3)).toContain(`<autoFilter ref="C1:D5"/>`);

    const workbookObj = parseXml(unzipped.workbookXml());
    const definedNames = workbookObj.workbook.definedNames.definedName;

    expect(definedNames).toHaveLength(2);
    expect(definedNames[0]["@_localSheetId"]).toBe("1");
    expect(definedNames[0]["#text"]).toBe("Sheet2!$A$1:$B$5");
    expect(definedNames[1]["@_localSheetId"]).toBe("2");
    expect(definedNames[1]["#text"]).toBe("Sheet3!$C$1:$D$5");
  });

  test("a sheet name that needs quoting is quoted in the defined name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("My Sheet");
    ws.setAutoFilter({ ref: "A1:B5" });

    const workbookObj = parseXml(
      unzipXlsx(wb.generateXlsxSync()).workbookXml()
    );

    expect(workbookObj.workbook.definedNames.definedName["#text"]).toBe(
      "'My Sheet'!$A$1:$B$5"
    );
  });

  test("no definedNames element is written when no sheet has an auto filter", () => {
    const wb = new Workbook();
    wb.addWorksheet("Sheet1");

    const workbookXml = unzipXlsx(wb.generateXlsxSync()).workbookXml();

    expect(workbookXml).not.toContain("<definedNames>");
  });

  test("WorkbookS does not support an auto filter", () => {
    const wb = new WorkbookS();
    const ws = wb.addWorksheet("Sheet1");

    expect(ws.autoFilter).toBeNull();
    expect(ws.autoFilterModule).toBeNull();

    const unzipped = unzipXlsx(wb.generateXlsxSync());

    expect(unzipped.sheetXml(1)).not.toContain("<autoFilter");
    expect(unzipped.workbookXml()).not.toContain("<definedNames>");
  });
});
