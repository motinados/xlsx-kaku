// TODO: Add internal targetMode
type WorksheetRel = {
  id: string;
  target: string;
  targetMode: "External" | null; // | "internal";
  relationshipType: string;
};

export class WorksheetRels {
  private rels: WorksheetRel[] = [];

  get relsLength(): number {
    return this.rels.length;
  }

  addWorksheetRel({
    target,
    targetMode,
    relationshipType,
  }: {
    target: string;
    targetMode: "External" | null;
    relationshipType: string;
  }): string {
    const id = "rId" + (this.rels.length + 1);
    const worksheetRel: WorksheetRel = {
      id,
      target,
      targetMode,
      relationshipType,
    };
    this.rels.push(worksheetRel);
    return id;
  }

  getWorksheetRels(): WorksheetRel[] {
    return this.rels;
  }

  // <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  //     <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://www.google.com" TargetMode="External"/>
  // </Relationships>
  makeXML(): string {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (const worksheetRel of this.rels) {
      const targetMode = worksheetRel.targetMode
        ? ` TargetMode="${worksheetRel.targetMode}"`
        : "";
      xml +=
        `<Relationship Id="${worksheetRel.id}"` +
        ` Type="${worksheetRel.relationshipType}"` +
        ` Target="${worksheetRel.target}"` +
        targetMode +
        "/>";
    }
    xml += "</Relationships>";
    return xml;
  }

  reset(): void {
    this.rels = [];
  }
}
