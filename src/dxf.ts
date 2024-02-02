import { Border, Font } from ".";

type Fill = {
  bgColor: string;
};

export type DxfStyle = {
  font: Font;
  fill: Fill;
  border: Border;
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
    let xml = `<dxfs count="${this.count}">`;
    this._dxf.forEach((style) => {
      xml += `<dxf><font>${makeFontXml(style.font)}</font><fill>${makeFillXml(
        style.fill
      )}</fill><border>${makeBorderXml(style.border)}</border></dxf>`;
    });
    xml += `</dxfs>`;
    return xml;
  }
}

export function makeFontXml(font: Font) {
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
  xml +=
    `<sz val="${font.size}"/>` +
    `<color rgb="${font.color}"/>` +
    `<name val="${font.name}"/>`;
  if (font.family) {
    xml += `<family val="${font.family}"/>`;
  }
  if (font.scheme) {
    xml += `<scheme val="${font.scheme}"/>`;
  }

  xml += "</font>";

  return xml;
}

export function makeFillXml(fill: Fill) {
  return `<patternFill><bgColor rgb="${fill.bgColor}"/></patternFill>`;
}

export function makeBorderXml(border: Border) {
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
