import { strFromU8, unzipSync } from "fflate";
import { Workbook, Worksheet } from "../src";
import { ImageStore } from "../src/imageStore";
import { genXlsx, genXlsxSync } from "../src/writer";
import { parseXml } from "./helper/helper";

describe("writer", () => {
  test("genXlsx", async () => {
    const imageStore = new ImageStore();
    const ws = new Worksheet("Sheet1", imageStore);
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = await genXlsx([ws], imageStore);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("genXlsxSync", () => {
    const imageStore = new ImageStore();
    const ws = new Worksheet("Sheet1", imageStore);
    ws.setCell(1, 1, { type: "string", value: "Hello" });
    const xlsx = genXlsxSync([ws], imageStore);
    expect(xlsx).toBeInstanceOf(Uint8Array);
  });

  test("A Error should occur when there is no sheet.", async () => {
    try {
      await genXlsx([], new ImageStore());
    } catch (e) {
      expect(e).toBeInstanceOf(Error);
    }
  });

  test("A Error should occur when there is no sheet.", () => {
    try {
      genXlsxSync([], new ImageStore());
    } catch (e) {
      expect(e).toBeInstanceOf(Error);
    }
  });

  test("keeps worksheet rels on sheet2 when only sheet2 has an external hyperlink", async () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");

    ws1.setCell(0, 0, { type: "string", value: "plain" });
    ws2.setCell(0, 0, {
      type: "hyperlink",
      text: "github",
      value: "https://github.com/",
      linkType: "external",
    });

    const xlsx = await wb.generateXlsx();
    const files = unzipSync(xlsx);

    expect(files["xl/worksheets/_rels/sheet1.xml.rels"]).toBeUndefined();
    expect(files["xl/worksheets/_rels/sheet2.xml.rels"]).toBeDefined();

    const sheet2Xml = strFromU8(files["xl/worksheets/sheet2.xml"]!);
    const sheet2Obj = parseXml(sheet2Xml);
    expect(sheet2Obj.worksheet.hyperlinks.hyperlink["@_r:id"]).toBe("rId1");

    const sheet2RelsXml = strFromU8(
      files["xl/worksheets/_rels/sheet2.xml.rels"]!
    );
    const sheet2RelsObj = parseXml(sheet2RelsXml);

    expect(sheet2RelsObj.Relationships.Relationship).toMatchObject({
      "@_Id": "rId1",
      "@_Type":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      "@_Target": "https://github.com/",
      "@_TargetMode": "External",
    });
  });

  test("keeps drawing files on sheet2 when only sheet2 has an image", async () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");

    ws1.setCell(0, 0, { type: "string", value: "plain" });
    await ws2.insertImage({
      displayName: "tiny",
      extension: "png",
      data: new Uint8Array([1, 2, 3, 4]),
      from: {
        col: 0,
        row: 0,
      },
      width: 10,
      height: 10,
    });

    const xlsx = await wb.generateXlsx();
    const files = unzipSync(xlsx);

    expect(files["xl/worksheets/_rels/sheet1.xml.rels"]).toBeUndefined();
    expect(files["xl/worksheets/_rels/sheet2.xml.rels"]).toBeDefined();
    expect(files["xl/drawings/drawing1.xml"]).toBeUndefined();
    expect(files["xl/drawings/_rels/drawing1.xml.rels"]).toBeUndefined();
    expect(files["xl/drawings/drawing2.xml"]).toBeDefined();
    expect(files["xl/drawings/_rels/drawing2.xml.rels"]).toBeDefined();

    const sheet2RelsXml = strFromU8(
      files["xl/worksheets/_rels/sheet2.xml.rels"]!
    );
    const sheet2RelsObj = parseXml(sheet2RelsXml);
    expect(sheet2RelsObj.Relationships.Relationship).toMatchObject({
      "@_Type":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
      "@_Target": "../drawings/drawing2.xml",
    });

    const drawing2RelsXml = strFromU8(
      files["xl/drawings/_rels/drawing2.xml.rels"]!
    );
    const drawing2RelsObj = parseXml(drawing2RelsXml);
    expect(drawing2RelsObj.Relationships.Relationship).toMatchObject({
      "@_Target": "../media/image1.png",
      "@_Type":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    });
  });
});
