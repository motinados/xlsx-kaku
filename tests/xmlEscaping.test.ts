import * as fflate from "fflate";
import { XMLValidator } from "fast-xml-parser";
import { Workbook } from "../src";
import { parseXml } from "./helper/helper";

const SHEET_NAME = `R&D <2024> "Q1"`;
const STRING_VALUE = `a & b < c > d "e" 'f'`;
const PADDED_STRING_VALUE = "  padded  ";
const FORMULA = `IF(A1<5,"low","high")&"!"`;
const URL = "https://example.com/search?a=1&b=2<c>";
const FORMAT_CODE = `0" & kg"`;
const CF_TEXT = `a & "b"`;

function toArray<T>(value: T | T[]): T[] {
  return Array.isArray(value) ? value : [value];
}

function generateXlsxParts(): Map<string, string> {
  const wb = new Workbook();
  const ws = wb.addWorksheet(SHEET_NAME);

  ws.setCell(0, 0, { type: "string", value: STRING_VALUE });
  ws.setCell(1, 0, { type: "string", value: PADDED_STRING_VALUE });
  ws.setCell(2, 0, { type: "formula", value: FORMULA });
  ws.setCell(3, 0, {
    type: "hyperlink",
    linkType: "external",
    text: STRING_VALUE,
    value: URL,
  });
  ws.setCell(4, 0, {
    type: "hyperlink",
    linkType: "internal",
    text: STRING_VALUE,
    value: `${SHEET_NAME}!A1`,
  });
  ws.setCell(5, 0, {
    type: "number",
    value: 1,
    style: { numberFormat: { formatCode: FORMAT_CODE } },
  });
  ws.setCell(6, 0, {
    type: "string",
    value: STRING_VALUE,
    style: { font: { name: `A&B "Gothic"` } },
  });
  ws.setConditionalFormatting({
    type: "containsText",
    sqref: "A1:A10",
    text: CF_TEXT,
    priority: 1,
    style: { font: { color: "FF9C0006" } },
  });
  ws.setAutoFilter({ ref: "A1:A7" });

  const unzipped = fflate.unzipSync(wb.generateXlsxSync());

  const decoder = new TextDecoder();
  const parts = new Map<string, string>();
  for (const [filename, content] of Object.entries(unzipped)) {
    if (filename.endsWith(".xml") || filename.endsWith(".rels")) {
      parts.set(filename, decoder.decode(content));
    }
  }
  return parts;
}

describe("xml escaping", () => {
  const parts = generateXlsxParts();

  test("every xml part is well-formed", () => {
    expect(parts.size).toBeGreaterThan(0);

    for (const [filename, xml] of parts) {
      const result = XMLValidator.validate(xml);
      // `validate` returns true, or an object describing the error.
      expect(result, `${filename} is not well-formed`).toBe(true);
    }
  });

  test("a string keeps its value", () => {
    const xml = parts.get("xl/sharedStrings.xml")!;
    const sharedStrings = parseXml(xml);
    const values = toArray(sharedStrings.sst.si).map(
      (si: { t: string | { "#text": string } }) =>
        typeof si.t === "object" ? si.t["#text"] : si.t
    );

    expect(values).toContain(STRING_VALUE);
    // The parser trims the text, so the raw xml is checked instead.
    expect(xml).toContain(
      `<si><t xml:space="preserve">${PADDED_STRING_VALUE}</t></si>`
    );
  });

  test("a formula keeps its value", () => {
    const worksheet = parseXml(parts.get("xl/worksheets/sheet1.xml")!);
    const cells = toArray<{ c: { f?: string } | { f?: string }[] }>(
      worksheet.worksheet.sheetData.row
    ).flatMap((row) => toArray(row.c));
    const formulaCell = cells.find((cell) => cell.f !== undefined);

    expect(formulaCell?.f).toBe(FORMULA);
  });

  test("a sheet name keeps its value", () => {
    const workbook = parseXml(parts.get("xl/workbook.xml")!);

    expect(workbook.workbook.sheets.sheet["@_name"]).toBe(SHEET_NAME);
  });

  test("a hyperlink keeps its url", () => {
    const rels = parseXml(parts.get("xl/worksheets/_rels/sheet1.xml.rels")!);
    const targets = toArray(rels.Relationships.Relationship).map(
      (rel: { "@_Target": string }) => rel["@_Target"]
    );

    expect(targets).toContain(URL);
  });

  test("a number format keeps its format code", () => {
    const styles = parseXml(parts.get("xl/styles.xml")!);
    const formatCodes = toArray(styles.styleSheet.numFmts.numFmt).map(
      (numFmt: { "@_formatCode": string }) => numFmt["@_formatCode"]
    );

    // A space in a format code is escaped with a backslash by Excel.
    expect(formatCodes).toContain(FORMAT_CODE.replace(/ /g, "\\ "));
  });

  test("the text of a conditional formatting keeps its value", () => {
    const worksheet = parseXml(parts.get("xl/worksheets/sheet1.xml")!);
    const cfRule = worksheet.worksheet.conditionalFormatting.cfRule;

    expect(cfRule["@_text"]).toBe(CF_TEXT);
    expect(cfRule.formula).toBe(`NOT(ISERROR(SEARCH("a & ""b""",A1)))`);
  });
});
