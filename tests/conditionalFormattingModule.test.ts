import { conditionalFormattingModule } from "../src/conditionalFormattingModule";
import { XlsxConditionalFormatting } from "../src/xml/worksheetXml";

describe("conditionalFormattingModule", () => {
  const module = conditionalFormattingModule();

  test("makeConditionalFormattingXml", () => {
    const conditionalFormattings: XlsxConditionalFormatting[] = [
      {
        sqref: "A1:A10",
        bottom: false,
        dxfId: 0,
        priority: 1,
        type: "top10",
        rank: 10,
        percent: true,
      },
    ];
    const actual = module.makeXmlElm(conditionalFormattings);
    const expected =
      `<conditionalFormatting sqref="A1:A10">` +
      `<cfRule type="top10" dxfId="0" priority="1" percent="1" rank="10"/>` +
      `</conditionalFormatting>`;
    expect(actual).toBe(expected);

    conditionalFormattings.push({
      sqref: "B1:B10",
      bottom: true,
      dxfId: 1,
      priority: 1,
      type: "top10",
      rank: 10,
      percent: false,
    });

    const acutual2 = module.makeXmlElm(conditionalFormattings);
    const expected2 =
      `<conditionalFormatting sqref="A1:A10">` +
      `<cfRule type="top10" dxfId="0" priority="1" percent="1" rank="10"/>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="B1:B10">` +
      `<cfRule type="top10" dxfId="1" priority="1" bottom="1" rank="10"/>` +
      `</conditionalFormatting>`;

    expect(acutual2).toBe(expected2);
  });

  test("makeConditionalFormattingXml with duplicateValues", () => {
    const conditionalFormattings: XlsxConditionalFormatting[] = [
      {
        sqref: "A1:A10",
        type: "duplicateValues",
        dxfId: 0,
        priority: 1,
      },
    ];
    const actual = module.makeXmlElm(conditionalFormattings);
    const expected =
      `<conditionalFormatting sqref="A1:A10">` +
      `<cfRule type="duplicateValues" dxfId="0" priority="1"/>` +
      `</conditionalFormatting>`;
    expect(actual).toBe(expected);
  });

  test("makeConditionalFormattingXml to compare numbers", () => {
    const conditionalFormattings: XlsxConditionalFormatting[] = [
      {
        sqref: "A1:A10",
        type: "cellIs",
        dxfId: 0,
        priority: 1,
        operator: "greaterThan",
        formula: "10",
      },
      {
        sqref: "B1:B10",
        type: "cellIs",
        dxfId: 1,
        priority: 2,
        operator: "lessThan",
        formula: "10",
      },
      {
        sqref: "C1:C10",
        type: "cellIs",
        dxfId: 2,
        priority: 3,
        operator: "equal",
        formula: "10",
      },
      {
        sqref: "D1:D10",
        type: "cellIs",
        dxfId: 3,
        priority: 4,
        operator: "between",
        formulaA: "3",
        formulaB: "8",
      },
    ];
    const actual = module.makeXmlElm(conditionalFormattings);
    const expected =
      `<conditionalFormatting sqref="A1:A10">` +
      `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">` +
      `<formula>10</formula>` +
      `</cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="B1:B10">` +
      `<cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan">` +
      `<formula>10</formula>` +
      `</cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="C1:C10">` +
      `<cfRule type="cellIs" dxfId="2" priority="3" operator="equal">` +
      `<formula>10</formula>` +
      `</cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="D1:D10">` +
      `<cfRule type="cellIs" dxfId="3" priority="4" operator="between">` +
      `<formula>3</formula>` +
      `<formula>8</formula>` +
      `</cfRule>` +
      `</conditionalFormatting>`;

    expect(actual).toBe(expected);
  });

  test("makeConditionalFormattingXml to compare strings", () => {
    const conditionalFormattings: XlsxConditionalFormatting[] = [
      {
        sqref: "A1:A1048576",
        type: "containsText",
        dxfId: 3,
        priority: 4,
        operator: "containsText",
        text: "a",
        formula: 'NOT(ISERROR(SEARCH("a",A1)))',
      },
      {
        sqref: "B1:B1048576",
        type: "notContainsText",
        dxfId: 2,
        priority: 3,
        operator: "notContains",
        text: "a",
        formula: 'ISERROR(SEARCH("a",B1))',
      },
      {
        sqref: "C1:C1048576",
        type: "beginsWith",
        dxfId: 1,
        priority: 2,
        operator: "beginsWith",
        text: "a",
        formula: 'LEFT(C1,LEN("a"))="a"',
      },
      {
        sqref: "D1:D1048576",
        type: "endsWith",
        dxfId: 0,
        priority: 1,
        operator: "endsWith",
        text: "a",
        formula: 'RIGHT(D1,LEN("a"))="a"',
      },
    ];

    const actual = module.makeXmlElm(conditionalFormattings);
    const expected =
      '<conditionalFormatting sqref="A1:A1048576">' +
      '<cfRule type="containsText" dxfId="3" priority="4" operator="containsText" text="a">' +
      '<formula>NOT(ISERROR(SEARCH("a",A1)))</formula>' +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="B1:B1048576">' +
      '<cfRule type="notContainsText" dxfId="2" priority="3" operator="notContains" text="a">' +
      '<formula>ISERROR(SEARCH("a",B1))</formula>' +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="C1:C1048576">' +
      '<cfRule type="beginsWith" dxfId="1" priority="2" operator="beginsWith" text="a">' +
      '<formula>LEFT(C1,LEN("a"))="a"</formula>' +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="D1:D1048576">' +
      '<cfRule type="endsWith" dxfId="0" priority="1" operator="endsWith" text="a">' +
      '<formula>RIGHT(D1,LEN("a"))="a"</formula>' +
      "</cfRule>" +
      "</conditionalFormatting>";
    expect(actual).toBe(expected);
  });

  test("makeConditionalFormattingXml to compare dates", () => {
    const conditionalFormattings: XlsxConditionalFormatting[] = [
      {
        sqref: "A1:A1048576",
        type: "timePeriod",
        dxfId: 0,
        priority: 1,
        timePeriod: "yesterday",
        formula: "FLOOR(A1,1)=TODAY()-1",
      },
      {
        sqref: "B1:B1048576",
        type: "timePeriod",
        dxfId: 9,
        priority: 10,
        timePeriod: "today",
        formula: "FLOOR(B1,1)=TODAY()",
      },
      {
        sqref: "C1:C10",
        type: "timePeriod",
        dxfId: 8,
        priority: 9,
        timePeriod: "tomorrow",
        formula: "FLOOR(C1,1)=TODAY()+1",
      },
      {
        sqref: "D1:D10",
        type: "timePeriod",
        dxfId: 7,
        priority: 8,
        timePeriod: "last7Days",
        formula: "AND(TODAY()-FLOOR(D1,1)<=6,FLOOR(D1,1)<=TODAY())",
      },
      {
        sqref: "E1:E10",
        type: "timePeriod",
        dxfId: 6,
        priority: 7,
        timePeriod: "lastWeek",
        formula:
          "AND(TODAY()-ROUNDDOWN(E1,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(E1,0)<(WEEKDAY(TODAY())+7))",
      },
      {
        sqref: "F1:F10",
        type: "timePeriod",
        dxfId: 5,
        priority: 6,
        timePeriod: "thisWeek",
        formula:
          "AND(TODAY()-ROUNDDOWN(F1,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(F1,0)-TODAY()<=7-WEEKDAY(TODAY()))",
      },
      {
        sqref: "G1:G10",
        type: "timePeriod",
        dxfId: 4,
        priority: 5,
        timePeriod: "nextWeek",
        formula:
          "AND(ROUNDDOWN(G1,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(G1,0)-TODAY()<(15-WEEKDAY(TODAY())))",
      },
      {
        sqref: "H1:H10",
        type: "timePeriod",
        dxfId: 3,
        priority: 4,
        timePeriod: "lastMonth",
        formula:
          "AND(MONTH(H1)=MONTH(EDATE(TODAY(),0-1)),YEAR(H1)=YEAR(EDATE(TODAY(),0-1)))",
      },
      {
        sqref: "I1:I10",
        type: "timePeriod",
        dxfId: 2,
        priority: 3,
        timePeriod: "thisMonth",
        formula: "AND(MONTH(I1)=MONTH(TODAY()),YEAR(I1)=YEAR(TODAY()))",
      },
      {
        sqref: "J1:J10",
        type: "timePeriod",
        dxfId: 1,
        priority: 2,
        timePeriod: "nextMonth",
        formula:
          "AND(MONTH(J1)=MONTH(EDATE(TODAY(),0+1)),YEAR(J1)=YEAR(EDATE(TODAY(),0+1)))",
      },
    ];
    const actual = module.makeXmlElm(conditionalFormattings);
    const expected =
      '<conditionalFormatting sqref="A1:A1048576">' +
      '<cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="yesterday">' +
      "<formula>FLOOR(A1,1)=TODAY()-1</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="B1:B1048576">' +
      '<cfRule type="timePeriod" dxfId="9" priority="10" timePeriod="today">' +
      "<formula>FLOOR(B1,1)=TODAY()</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="C1:C10">' +
      '<cfRule type="timePeriod" dxfId="8" priority="9" timePeriod="tomorrow">' +
      "<formula>FLOOR(C1,1)=TODAY()+1</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="D1:D10">' +
      '<cfRule type="timePeriod" dxfId="7" priority="8" timePeriod="last7Days">' +
      "<formula>AND(TODAY()-FLOOR(D1,1)<=6,FLOOR(D1,1)<=TODAY())</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="E1:E10">' +
      '<cfRule type="timePeriod" dxfId="6" priority="7" timePeriod="lastWeek">' +
      "<formula>AND(TODAY()-ROUNDDOWN(E1,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(E1,0)<(WEEKDAY(TODAY())+7))</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="F1:F10">' +
      '<cfRule type="timePeriod" dxfId="5" priority="6" timePeriod="thisWeek">' +
      "<formula>AND(TODAY()-ROUNDDOWN(F1,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(F1,0)-TODAY()<=7-WEEKDAY(TODAY()))</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="G1:G10">' +
      '<cfRule type="timePeriod" dxfId="4" priority="5" timePeriod="nextWeek">' +
      "<formula>AND(ROUNDDOWN(G1,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(G1,0)-TODAY()<(15-WEEKDAY(TODAY())))</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="H1:H10">' +
      '<cfRule type="timePeriod" dxfId="3" priority="4" timePeriod="lastMonth">' +
      "<formula>AND(MONTH(H1)=MONTH(EDATE(TODAY(),0-1)),YEAR(H1)=YEAR(EDATE(TODAY(),0-1)))</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="I1:I10">' +
      '<cfRule type="timePeriod" dxfId="2" priority="3" timePeriod="thisMonth">' +
      "<formula>AND(MONTH(I1)=MONTH(TODAY()),YEAR(I1)=YEAR(TODAY()))</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>" +
      '<conditionalFormatting sqref="J1:J10">' +
      '<cfRule type="timePeriod" dxfId="1" priority="2" timePeriod="nextMonth">' +
      "<formula>AND(MONTH(J1)=MONTH(EDATE(TODAY(),0+1)),YEAR(J1)=YEAR(EDATE(TODAY(),0+1)))</formula>" +
      "</cfRule>" +
      "</conditionalFormatting>";
    expect(actual).toBe(expected);
  });
});
