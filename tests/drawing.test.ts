import { DrawingRels } from "../src/drawingRels";

describe("DrawingRels", () => {
  test("addDrawingRel should add a new drawing relationship", () => {
    const drawingRels = new DrawingRels();
    const target = "../media/image1.png";
    const id = drawingRels.addDrawingRel({
      target: target,
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    });
    const rels = drawingRels.rels;
    expect(rels.length).toBe(1);
    expect(rels[0]).toBeDefined();
    expect(rels[0]!.id).toBe(id);
    expect(rels[0]!.target).toBe(target);
    expect(rels[0]!.relationshipType).toBe(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    );
  });
  test("makeXml should generate the correct XML string", () => {
    const drawingRels = new DrawingRels();
    const target1 = "../media/image1.png";
    drawingRels.addDrawingRel({
      target: target1,
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    });

    const xml = drawingRels.makeXml();
    expect(xml).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
        "</Relationships>"
    );
  });

  test("reset should reset the relationships", () => {
    const drawingRels = new DrawingRels();
    drawingRels.addDrawingRel({
      target: "../media/image1.png",
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    });
    drawingRels.reset();
    expect(drawingRels.rels.length).toBe(0);
  });
});
