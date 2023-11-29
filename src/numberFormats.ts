export class NumberFormats {
  // custom numFmtId starts from 176
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
    ["mm-dd-yy", 14],
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

    // custom numFmtId starts from 176
    if (this.lastCustomNumFmtId === -1) {
      this.lastCustomNumFmtId = 176;
    }
    const numFmtId = this.lastCustomNumFmtId;
    this.numFmts.set(formatCode, numFmtId);
    this.lastCustomNumFmtId++;
    return numFmtId;
  }
}
