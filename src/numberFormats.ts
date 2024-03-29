export type NumberFormat = {
  formatCode: string;
};

export class NumberFormats {
  // custom numFmtId starts from 164
  private lastCustomNumFmtId = -1;
  private numFmts = new Map([
    ["0", 1],
    ["0.00", 2],
    ["#,##0", 3],
    ["#,##0.00", 4],
    ["0%", 9],
    ["0.00%", 10],
    ["0.00E+00", 11],
    ["# ?/?", 12],
    ["# ??/??", 13],
    // ["mm-dd-yy", 14],
    ["d-mmm-yy", 15],
    ["d-mmm", 16],
    ["mmm-yy", 17],
    ["h:mm AM/PM", 18],
    ["h:mm:ss AM/PM", 19],
    ["h:mm", 20],
    ["h:mm:ss", 21],
    ["m/d/yy h:mm", 22],
    ["#,##0 ;(#,##0)", 37],
    ["#,##0 ;[Red](#,##0)", 38],
    ["#,##0.00;(#,##0.00)", 39],
    ["#,##0.00;[Red](#,##0.00)", 40],
    ["mm:ss", 45],
    ["[h]:mm:ss", 46],
    ["mmss.0", 47],
    ["##0.0E+0", 48],
    ["@", 49],
  ]);

  getNumFmtId(formatCode: string): number {
    const id = this.numFmts.get(formatCode);
    if (id !== undefined) {
      return id;
    }

    // custom numFmtId starts from 164
    if (this.lastCustomNumFmtId === -1) {
      this.lastCustomNumFmtId = 164;
    }
    const numFmtId = this.lastCustomNumFmtId;
    this.numFmts.set(formatCode, numFmtId);
    this.lastCustomNumFmtId++;
    return numFmtId;
  }

  // items 164 and above are custom numFmts
  private extractItemsWithIdAbove176(): Map<string, number> {
    const items = new Map<string, number>();
    this.numFmts.forEach((numFmtId, formatCode) => {
      if (numFmtId <= 163) {
        return;
      }
      items.set(formatCode, numFmtId);
    });
    return items;
  }

  // numFmt smaller than 163 is provided by default, so threre is no need to make it xml.
  // <numFmts count="1">
  //   <numFmt numFmtId="164" formatCode="yyyy/m/d\ h:mm"/>
  // </numFmts>
  makeXml(): string {
    function escapeString(input: string): string {
      const regex = /(\[\$-\d+\])([^\[]+)/g;

      if (input.match(regex)) {
        return input.replace(regex, (_, p1, p2) => {
          return p1 + p2.replace(/-/g, "\\-").replace(/ /g, "\\ ");
        });
      } else {
        return input.replace(/-/g, "\\-").replace(/ /g, "\\ ");
      }
    }

    const items = this.extractItemsWithIdAbove176();
    if (items.size === 0) {
      return "";
    }

    let xml = `<numFmts count="${items.size}">`;
    items.forEach((numFmtId, formatCode) => {
      const code = escapeString(formatCode);
      xml += `<numFmt numFmtId="${numFmtId}" formatCode="${code}"/>`;
    });
    xml += "</numFmts>";
    return xml;
  }
}
