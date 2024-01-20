import { WorksheetRels } from "../src/worksheetRels";

describe("WorksheetRels", () => {
  test("addWorksheetRel should add a new worksheet relationship", () => {
    const worksheetRels = new WorksheetRels();
    const target = "http://www.google.com";
    const id = worksheetRels.addWorksheetRel(target);
    const rels = worksheetRels.getWorksheetRels();

    expect(rels.length).toBe(1);
    expect(rels[0]).toBeDefined();
    expect(rels[0]!.id).toBe(id);
    expect(rels[0]!.target).toBe(target);
    expect(rels[0]!.targetMode).toBe("External");
  });

  test("getWorksheetRels should return all worksheet relationships", () => {
    const worksheetRels = new WorksheetRels();

    const target1 = "http://www.google.com";
    const target2 = "http://www.github.com";

    worksheetRels.addWorksheetRel(target1);
    worksheetRels.addWorksheetRel(target2);

    const rels = worksheetRels.getWorksheetRels();

    expect(rels.length).toBe(2);
    expect(rels[0]).toBeDefined();
    expect(rels[1]).toBeDefined();
    expect(rels[0]!.target).toBe(target1);
    expect(rels[1]!.target).toBe(target2);
  });

  test("makeXML should generate the correct XML string", () => {
    const worksheetRels = new WorksheetRels();

    const target1 = "http://www.google.com";
    const target2 = "http://www.github.com";

    worksheetRels.addWorksheetRel(target1);
    worksheetRels.addWorksheetRel(target2);

    const expectedXML =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://www.google.com" TargetMode="External"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://www.github.com" TargetMode="External"/>' +
      "</Relationships>";

    const xml = worksheetRels.makeXML();

    expect(xml).toBe(expectedXML);
  });

  test("reset", () => {
    const worksheetRels = new WorksheetRels();

    const target1 = "http://www.google.com";
    const target2 = "http://www.github.com";

    worksheetRels.addWorksheetRel(target1);
    worksheetRels.addWorksheetRel(target2);

    worksheetRels.reset();

    const rels = worksheetRels.getWorksheetRels();

    expect(rels).toEqual([]);
  });
});
