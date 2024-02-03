import { Border, Font } from ".";

type Fill = {
  bgColor: string;
};

export type DxfStyle = {
  font?: Font;
  fill?: Fill;
  border?: Border;
};

// TODO: integrate styles
export class Dxf {
  private _dxf: Map<number, DxfStyle> = new Map();
  get count(): number {
    return this._dxf.size;
  }

  addStyle(style: DxfStyle) {
    let key = this._dxf.size;
    this._dxf.set(key, style);
    return key;
  }

  makeXml(): string {
    if (this.count === 0) {
      return "";
    }

    let xml = `<dxfs count="${this.count}">`;
    this._dxf.forEach((style) => {
      xml +=
        "<dxf>" +
        makeFontXml(style.font) +
        makeFillXml(style.fill) +
        makeBorderXml(style.border) +
        "</dxf>";
    });
    xml += `</dxfs>`;
    return xml;
  }
}

export function makeFontXml(font: Font | undefined) {
  if (!font) {
    return "";
  }

  let xml = "<font>";

  if (font.bold) {
    xml += `<b/>`;
  }

  if (font.italic) {
    xml += `<i/>`;
  }

  if (font.strike) {
    xml += `<strike/>`;
  }

  if (font.underline) {
    if (font.underline == "double") {
      xml += `<u val="double"/>`;
    } else {
      xml += `<u/>`;
    }
  }

  if (font.size) {
    xml += `<sz val="${font.size}"/>`;
  }

  if (font.color) {
    xml += `<color rgb="${font.color}"/>`;
  }

  if (font.name) {
    xml += `<name val="${font.name}"/>`;
  }

  if (font.family) {
    xml += `<family val="${font.family}"/>`;
  }
  if (font.scheme) {
    xml += `<scheme val="${font.scheme}"/>`;
  }

  xml += "</font>";

  return xml;
}

export function makeFillXml(fill: Fill | undefined) {
  if (!fill) {
    return "";
  }

  return `<fill><patternFill><bgColor rgb="${fill.bgColor}"/></patternFill></fill>`;
}

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
