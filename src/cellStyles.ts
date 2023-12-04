import { stringifySorted } from "./fonts";

// <cellStyles count="2">
//     <cellStyle name="Hyperlink" xfId="1" xr:uid="{00000000-000B-0000-0000-000008000000}"/>
//     <cellStyle name="標準" xfId="0" builtinId="0"/>
// </cellStyles>
type CellStyle =
  | { name: string; xfId: number; builtinId: number }
  | { name: string; xfId: number; uid: string };

export class CellStyles {
  private cellStyles = new Map<string, number>([
    [
      stringifySorted({
        name: "標準",
        xfId: 0,
        builtinId: 0,
      }),
      0,
    ],
  ]);

  getCellStyleId(cellStyle: CellStyle): number {
    const key = stringifySorted(cellStyle);
    const id = this.cellStyles.get(key);
    if (id !== undefined) {
      return id;
    }

    const cellStyleId = this.cellStyles.size;
    this.cellStyles.set(key, cellStyleId);
    return cellStyleId;
  }

  makeXml(): string {
    let xml = `<cellStyles count="${this.cellStyles.size}">`;
    this.cellStyles.forEach((_, key) => {
      const cellStyle = JSON.parse(key) as CellStyle;
      xml += `<cellStyle name="${cellStyle.name}" xfId="${cellStyle.xfId}"`;
      if ("uid" in cellStyle) {
        xml += ` xr:uid="${cellStyle.uid}"`;
      } else {
        xml += ` builtinId="${cellStyle.builtinId}"`;
      }
      xml += "/>";
    });
    xml += "</cellStyles>";
    return xml;
  }
}
