import { stringifySorted } from "./utils";

export type Alignment = {
  horizontal?: "left" | "center" | "right";
  vertical?: "top" | "center" | "bottom";
  textRotation?: number;
  wordwrap?: boolean;
};

export type CellXf = {
  fillId: number;
  fontId: number;
  borderId: number;
  numFmtId: number;
  xfId?: number;
  alignment?: Alignment;
};

export class CellXfs {
  private _cellXfs = new Map<string, number>([
    [
      stringifySorted({
        fillId: 0,
        fontId: 0,
        borderId: 0,
        numFmtId: 0,
      }),
      0,
    ],
  ]);

  get count(): number {
    return this._cellXfs.size;
  }

  get cellXfs(): Map<string, number> {
    return this._cellXfs;
  }

  getCellXfId(cellXf: CellXf): number {
    const key = stringifySorted(cellXf);
    const id = this._cellXfs.get(key);
    if (id !== undefined) {
      return id;
    }

    const cellXfId = this.count;
    this._cellXfs.set(key, cellXfId);
    return cellXfId;
  }

  // <cellXfs count="2">
  //   <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  //   <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
  // </cellXfs>
  makeXml(): string {
    let xml = `<cellXfs count="${this.count}">`;
    this._cellXfs.forEach((_, key) => {
      const cellXf = JSON.parse(key) as CellXf;
      const xfId = cellXf.xfId ?? 0;

      xml += `<xf numFmtId="${cellXf.numFmtId}" fontId="${cellXf.fontId}" fillId="${cellXf.fillId}" borderId="${cellXf.borderId}" xfId="${xfId}"`;

      if (cellXf.alignment) {
        xml += ' applyAlignment="1"';
      }
      if (cellXf.fillId > 0) {
        xml += ' applyFill="1"';
      }
      if (cellXf.numFmtId > 0) {
        xml += ' applyNumberFormat="1"';
      }
      if (cellXf.fontId > 0) {
        xml += ' applyFont="1"';
      }
      if (cellXf.borderId > 0) {
        xml += ' applyBorder="1"';
      }

      if (cellXf.alignment) {
        xml += ">";
        xml += "<alignment";
        if (cellXf.alignment.horizontal) {
          xml += ` horizontal="${cellXf.alignment.horizontal}"`;
        }
        if (cellXf.alignment.vertical) {
          xml += ` vertical="${cellXf.alignment.vertical}"`;
        }
        if (cellXf.alignment.textRotation) {
          xml += ` textRotation="${cellXf.alignment.textRotation}"`;
        }
        if (cellXf.alignment.wordwrap) {
          xml += ` wrapText="1"`;
        }
        xml += "/>";
        xml += "</xf>";
      } else {
        xml += "/>";
      }
    });
    xml += "</cellXfs>";
    return xml;
  }
}
