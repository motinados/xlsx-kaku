import { makeFontXml } from "./dxf";
import { stringifySorted } from "./utils";

export type Font = {
  name?: string;
  size?: number;
  // TODO: support theme color
  color?: string;
  family?: number;
  scheme?: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  /**
   * "single": single underline, "double": double underline
   */
  underline?: "single" | "double";
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
      xml += makeFontXml(font);
    });
    xml += "</fonts>";

    return xml;
  }
}
