import { hasSheetName } from "./utils";

type Hyperlink =
  | {
      linkType: "external" | "email";
      ref: string;
      rid: string;
      uuid: string;
    }
  | {
      linkType: "internal";
      ref: string;
      location: string;
      display: string;
      uuid: string;
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
      if (hyperlink.linkType === "external" || hyperlink.linkType === "email") {
        xml +=
          '<hyperlink ref="' +
          hyperlink.ref +
          '" r:id="' +
          hyperlink.rid +
          '" xr:uid="{' +
          hyperlink.uuid +
          '}"/>';
      } else if (hyperlink.linkType === "internal") {
        let location;
        if (hasSheetName(hyperlink.location)) {
          const [sheetName, cellAddress] = hyperlink.location.split("!");
          location = `'${sheetName}'!${cellAddress}`;
        } else {
          location = hyperlink.location;
        }

        xml +=
          '<hyperlink ref="' +
          hyperlink.ref +
          '" location="' +
          location +
          '" display="' +
          hyperlink.display +
          '" xr:uid="{' +
          hyperlink.uuid +
          '}"/>';
      }
    }
    xml += "</hyperlinks>";
    return xml;
  }
}
