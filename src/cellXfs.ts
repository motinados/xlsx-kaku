import { stringifySorted } from "./fonts";

type CellXf = {
  fillId: number;
  fontId: number;
  borderId: number;
  numFmtId: number;
  xfId?: number;
};

export class CellXfs {
  private cellXfs = new Map<string, number>([
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

  getCellXfId(cellXf: CellXf): number {
    const key = stringifySorted(cellXf);
    const id = this.cellXfs.get(key);
    if (id !== undefined) {
      return id;
    }

    const cellXfId = this.cellXfs.size;
    this.cellXfs.set(key, cellXfId);
    return cellXfId;
  }

  // <cellXfs count="2">
  //   <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  //   <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
  // </cellXfs>
  makeXml(): string {
    let xml = `<cellXfs count="${this.cellXfs.size}">`;
    this.cellXfs.forEach((_, key) => {
      const cellXf = JSON.parse(key) as CellXf;
      const xfId = cellXf.xfId ?? 0;
      xml += `<xf numFmtId="${cellXf.numFmtId}" fontId="${cellXf.fontId}" fillId="${cellXf.fillId}" borderId="${cellXf.borderId}" xfId="${xfId}"`;
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
      xml += "/>";
    });
    xml += "</cellXfs>";
    return xml;
  }
}
