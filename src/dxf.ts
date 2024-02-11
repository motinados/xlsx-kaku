import { Border, Font } from ".";
import { makeBorderXml } from "./borders";
import { Fill, makeFillXml } from "./fills";
import { makeFontXml } from "./fonts";

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
