import { stringifySorted } from "./utils";

// https://learn.microsoft.com/ja-jp/dotnet/api/documentformat.openxml.spreadsheet.borderstylevalues?view=openxml-2.8.1
type BorderStyle =
  | "dashDot"
  | "dashDotDot"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "medium"
  | "mediumDashDot"
  | "mediumDashDotDot"
  | "mediumDashed"
  | "none"
  | "slantDashDot"
  | "thick"
  | "thin";

export type Border = {
  left?: {
    style: BorderStyle;
    color: string;
  };
  right?: {
    style: BorderStyle;
    color: string;
  };
  top?: {
    style: BorderStyle;
    color: string;
  };
  bottom?: {
    style: BorderStyle;
    color: string;
  };
  diagonal?: {
    style: BorderStyle;
    color: string;
  };
};

export function makeBorderXml(border: Border | undefined) {
  if (!border) {
    return "";
  }

  let xml = "<border>";
  if (border.left) {
    xml +=
      `<left style="${border.left.style}">` +
      `<color rgb="${border.left.color}"/>` +
      "</left>";
  } else {
    xml += "<left/>";
  }
  if (border.right) {
    xml +=
      `<right style="${border.right.style}">` +
      `<color rgb="${border.right.color}"/>` +
      "</right>";
  } else {
    xml += "<right/>";
  }
  if (border.top) {
    xml +=
      `<top style="${border.top.style}">` +
      `<color rgb="${border.top.color}"/>` +
      "</top>";
  } else {
    xml += "<top/>";
  }
  if (border.bottom) {
    xml +=
      `<bottom style="${border.bottom.style}">` +
      `<color rgb="${border.bottom.color}"/>` +
      "</bottom>";
  } else {
    xml += "<bottom/>";
  }
  if (border.diagonal) {
    xml +=
      `<diagonal style="${border.diagonal.style}">` +
      `<color rgb="${border.diagonal.color}"/>` +
      "</diagonal>";
  } else {
    xml += "<diagonal/>";
  }
  xml += "</border>";

  return xml;
}

export class Borders {
  // default border
  //   <borders count="1">
  //       <border>
  //           <left/>
  //           <right/>
  //           <top/>
  //           <bottom/>
  //           <diagonal/>
  //       </border>
  //   </borders>
  private borders = new Map<string, number>();

  // borders map does not have a default border, so retuns size + 1
  get bordersSize() {
    return this.borders.size + 1;
  }

  getBorderId(border: Border): number {
    if (Object.keys(border).length === 0) {
      return 0;
    }

    const key = stringifySorted(border);
    const id = this.borders.get(key);
    if (id !== undefined) {
      return id;
    }

    const borderId = this.bordersSize;
    this.borders.set(key, borderId);
    return borderId;
  }

  // <borders count="2">
  //       <border>
  //           <left/>
  //           <right/>
  //           <top/>
  //           <bottom/>
  //           <diagonal/>
  //       </border>
  //       <border>
  //           <left style="thin">
  //               <color rgb="FF000000"/>
  //           </left>
  //           <right style="thin">
  //               <color rgb="FF000000"/>
  //           </right>
  //           <top style="thin">
  //               <color rgb="FF000000"/>
  //           </top>
  //           <bottom style="thin">
  //               <color rgb="FF000000"/>
  //           </bottom>
  //           <diagonal/>
  //       </border>
  //   </borders>
  makeXml(): string {
    let xml = `<borders count="${this.bordersSize}">`;

    // default border
    xml +=
      "<border>" +
      "<left/>" +
      "<right/>" +
      "<top/>" +
      "<bottom/>" +
      "<diagonal/>" +
      "</border>";

    this.borders.forEach((_, key) => {
      const border = JSON.parse(key) as Border;
      xml += makeBorderXml(border);
    });
    xml += "</borders>";

    return xml;
  }
}
