import { Styles } from "../src/styles";
describe("Styles", () => {
  test("default getNumFmtId", () => {
    const styles = new Styles();
    expect(styles.getNumFmtId("0")).toBe(1);
    expect(styles.getNumFmtId("0.00")).toBe(2);
    expect(styles.getNumFmtId("#,##0")).toBe(3);
    expect(styles.getNumFmtId("#,##0.00")).toBe(4);
    expect(styles.getNumFmtId("0%")).toBe(9);
    expect(styles.getNumFmtId("0.00%")).toBe(10);
    expect(styles.getNumFmtId("0.00E+00")).toBe(11);
    expect(styles.getNumFmtId("# ?/?")).toBe(12);
    expect(styles.getNumFmtId("# ??/??")).toBe(13);
    expect(styles.getNumFmtId("mm-dd-yy")).toBe(14);
    expect(styles.getNumFmtId("d-mmm-yy")).toBe(15);
    expect(styles.getNumFmtId("d-mmm")).toBe(16);
    expect(styles.getNumFmtId("mmm-yy")).toBe(17);
    expect(styles.getNumFmtId("h:mm AM/PM")).toBe(18);
    expect(styles.getNumFmtId("h:mm:ss AM/PM")).toBe(19);
    expect(styles.getNumFmtId("h:mm")).toBe(20);
    expect(styles.getNumFmtId("h:mm:ss")).toBe(21);
    expect(styles.getNumFmtId("m/d/yy h:mm")).toBe(22);
    expect(styles.getNumFmtId("#,##0 ;(#,##0)")).toBe(37);
    expect(styles.getNumFmtId("#,##0 ;[Red](#,##0)")).toBe(38);
    expect(styles.getNumFmtId("#,##0.00;(#,##0.00)")).toBe(39);
    expect(styles.getNumFmtId("#,##0.00;[Red](#,##0.00)")).toBe(40);
    expect(styles.getNumFmtId("mm:ss")).toBe(45);
    expect(styles.getNumFmtId("[h]:mm:ss")).toBe(46);
    expect(styles.getNumFmtId("mmss.0")).toBe(47);
    expect(styles.getNumFmtId("##0.0E+0")).toBe(48);
    expect(styles.getNumFmtId("@")).toBe(49);
  });

  test("custom getNumFmtId", () => {
    const styles = new Styles();
    expect(styles.getNumFmtId("yyyy-mm-dd")).toBe(176);
    expect(styles.getNumFmtId("mm-yyyy-dd")).toBe(177);
  });
});
