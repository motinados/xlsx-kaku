import { stringifySorted } from "./utils";

type CellStyleXf = {
  fillId: number;
  fontId: number;
  borderId: number;
  numFmtId: number;
};

export class CellStyleXfs {
  private cellStyleXfs = new Map<string, number>([
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

  getCellStyleXfId(cellStyleXf: CellStyleXf): number {
    const key = stringifySorted(cellStyleXf);
    const id = this.cellStyleXfs.get(key);
    if (id !== undefined) {
      return id;
    }

    const cellStyleXfId = this.cellStyleXfs.size;
    this.cellStyleXfs.set(key, cellStyleXfId);
    return cellStyleXfId;
  }

  makeXml(): string {
    let xml = `<cellStyleXfs count="${this.cellStyleXfs.size}">`;
    this.cellStyleXfs.forEach((_, key) => {
      const cellStyleXf = JSON.parse(key) as CellStyleXf;
      xml += `<xf numFmtId="${cellStyleXf.numFmtId}" fontId="${cellStyleXf.fontId}" fillId="${cellStyleXf.fillId}" borderId="${cellStyleXf.borderId}" xfId="0"`;
      if (cellStyleXf.fillId > 0) {
        xml += ' applyFill="1"';
      }
      if (cellStyleXf.numFmtId > 0) {
        xml += ' applyNumberFormat="1"';
      }
      if (cellStyleXf.fontId > 0) {
        xml += ' applyFont="1"';
      }
      if (cellStyleXf.borderId > 0) {
        xml += ' applyBorder="1"';
      }
      xml += "/>";
    });
    xml += "</cellStyleXfs>";
    return xml;
  }
}
