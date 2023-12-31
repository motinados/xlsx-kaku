import { stringifySorted } from "./utils";

export type Font = {
  name: string;
  size: number;
  // TODO: support theme color
  color: string;
  family?: number;
  scheme?: string;
  underline?: boolean;
};

export class Fonts {
  private fonts = new Map<string, number>([
    [
      stringifySorted({
        name: "Calibri",
        size: 11,
        color: "000000",
      }),
      0,
    ],
  ]);

  getFontId(font: Font): number {
    const key = stringifySorted(font);
    const id = this.fonts.get(key);
    if (id !== undefined) {
      return id;
    }

    const fontId = this.fonts.size;
    this.fonts.set(key, fontId);
    return fontId;
  }

  // <font>
  //   <sz val="11"/>
  //   <color theme="1"/>
  //   <name val="Calibri"/>
  //   <family val="2"/>
  //   <scheme val="minor"/>
  // </font>
  makeXml(): string {
    let xml = `<fonts count="${this.fonts.size}">`;

    this.fonts.forEach((_, key) => {
      const font = JSON.parse(key) as Font;
      xml += "<font>";
      if (font.underline) {
        xml += `<u/>`;
      }
      xml += `<sz val="${font.size}"/>`;
      xml += `<color rgb="${font.color}"/>`;
      xml += `<name val="${font.name}"/>`;
      if (font.family) {
        xml += `<family val="${font.family}"/>`;
      }
      if (font.scheme) {
        xml += `<scheme val="${font.scheme}"/>`;
      }

      xml += "</font>";
    });
    xml += "</fonts>";

    return xml;
  }
}
