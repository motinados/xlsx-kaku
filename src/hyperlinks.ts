// TODO: Add internal hyperlinks
type Hyperlink = {
  ref: string;
  uuid: string;
  targetMode: "external"; // | "internal";
};

export class Hyperlinks {
  private hyperlinks: Hyperlink[] = [];

  addHyperlink(hyperlink: Hyperlink): void {
    this.hyperlinks.push(hyperlink);
  }

  getHyperlinks(): Hyperlink[] {
    return this.hyperlinks;
  }

  // <hyperlinks>
  //     <hyperlink ref="A1" r:id="rId1" xr:uid="{9F39A655-5247-453A-B80C-23E93683E969}"/>
  // </hyperlinks>
  makeXML(): string {
    let xml = "";
    xml += "<hyperlinks>";
    for (const hyperlink of this.hyperlinks) {
      xml +=
        '<hyperlink ref="' +
        hyperlink.ref +
        '" r:id="rId1' +
        '" xr:uid="{' +
        hyperlink.uuid +
        '}"/>';
    }
    xml += "</hyperlinks>";
    return xml;
  }
}
