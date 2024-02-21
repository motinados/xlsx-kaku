type DrawingRel = {
  id: string;
  target: string;
  relationshipType: string;
};

export class DrawingRels {
  private _rels: DrawingRel[] = [];
  constructor() {}
  get rels(): DrawingRel[] {
    return this._rels;
  }

  get length(): number {
    return this._rels.length;
  }

  addDrawingRel({
    target,
    relationshipType,
  }: {
    target: string;
    relationshipType: string;
  }): string {
    const id = "rId" + (this._rels.length + 1);
    const drawingRel: DrawingRel = {
      id,
      target,
      relationshipType,
    };
    this._rels.push(drawingRel);
    return id;
  }
  /**
   * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
   */
  makeXml(): string {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

    for (const worksheetRel of this.rels) {
      xml +=
        `<Relationship Id="${worksheetRel.id}"` +
        ` Type="${worksheetRel.relationshipType}"` +
        ` Target="${worksheetRel.target}"` +
        "/>";
    }

    xml += "</Relationships>";
    return xml;
  }

  reset() {
    this._rels = [];
  }
}
