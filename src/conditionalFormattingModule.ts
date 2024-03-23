import { v4 as uuidv4 } from "uuid";
import { Dxf } from "./dxf";
import { getFirstAddress } from "./utils";
import { ConditionalFormatting } from "./worksheet";
import { XlsxConditionalFormatting } from "./xml/worksheetXml";

export type ConditionalFormattingModule = {
  name: string;
  getConditionalFormattings(): ConditionalFormatting[];
  add(conditionalFormatting: ConditionalFormatting): void;
  createXlsxConditionalFormatting(
    conditionalFormattings: ConditionalFormatting[],
    dxf: Dxf
  ): XlsxConditionalFormatting[];
  makeXmlElm(formattings: XlsxConditionalFormatting[]): string;
};

export function conditionalFormattingModule(): ConditionalFormattingModule {
  const conditionalFormattings: ConditionalFormatting[] = [];
  return {
    name: "conditional-formatting",
    getConditionalFormattings() {
      return conditionalFormattings;
    },
    add(conditionalFormatting: ConditionalFormatting) {
      conditionalFormattings.push(conditionalFormatting);
    },
    createXlsxConditionalFormatting(
      conditionalFormattings: ConditionalFormatting[],
      dxf: Dxf
    ) {
      const xcfs: XlsxConditionalFormatting[] = [];
      if (conditionalFormattings.length > 0) {
        for (const cf of conditionalFormattings) {
          if (cf.type === "dataBar") {
            const conditionalFormatting: XlsxConditionalFormatting = {
              type: "dataBar",
              sqref: cf.sqref,
              priority: cf.priority,
              color: cf.color,
              x14Id: uuidv4(),
              border: cf.border,
              gradient: cf.gradient,
              negativeBarBorderColorSameAsPositive:
                cf.negativeBarBorderColorSameAsPositive,
            };
            xcfs.push(conditionalFormatting);
            continue;
          } else if (cf.type === "colorScale") {
            const conditionalFormatting: XlsxConditionalFormatting = {
              type: "colorScale",
              sqref: cf.sqref,
              priority: cf.priority,
              colorScale: cf.colorScale,
            };
            xcfs.push(conditionalFormatting);
            continue;
          } else if (cf.type === "iconSet") {
            const conditionalFormatting: XlsxConditionalFormatting = {
              type: "iconSet",
              sqref: cf.sqref,
              priority: cf.priority,
              iconSet: cf.iconSet,
            };
            xcfs.push(conditionalFormatting);
            continue;
          }

          const id = dxf.addStyle(cf.style);

          switch (cf.type) {
            case "top":
            case "bottom": {
              const bottom = cf.type === "bottom";
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "top10",
                sqref: cf.sqref,
                priority: cf.priority,
                percent: cf.percent,
                bottom,
                rank: cf.rank,
                dxfId: id,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "aboveAverage":
            case "belowAverage":
            case "atOrAboveAverage":
            case "atOrBelowAverage": {
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "aboveAverage",
                sqref: cf.sqref,
                priority: cf.priority,
                aboveAverage:
                  cf.type === "aboveAverage" || cf.type === "atOrAboveAverage",
                equalAverage:
                  cf.type === "atOrAboveAverage" ||
                  cf.type === "atOrBelowAverage",
                dxfId: id,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "duplicateValues": {
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "duplicateValues",
                sqref: cf.sqref,
                priority: cf.priority,
                dxfId: id,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "greaterThan":
            case "lessThan":
            case "equal": {
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "cellIs",
                sqref: cf.sqref,
                priority: cf.priority,
                operator: cf.type,
                formula: "" + cf.formula,
                dxfId: id,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "between": {
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "cellIs",
                sqref: cf.sqref,
                priority: cf.priority,
                operator: "between",
                formulaA: "" + cf.formulaA,
                formulaB: "" + cf.formulaB,
                dxfId: id,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "containsText": {
              const firstCell = getFirstAddress(cf.sqref);
              const formula = `NOT(ISERROR(SEARCH("${cf.text}",${firstCell})))`;
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "containsText",
                sqref: cf.sqref,
                priority: cf.priority,
                operator: "containsText",
                text: cf.text,
                dxfId: id,
                formula: formula,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "notContainsText": {
              const firstCell = getFirstAddress(cf.sqref);
              const formula = `ISERROR(SEARCH("${cf.text}",${firstCell}))`;
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "notContainsText",
                sqref: cf.sqref,
                priority: cf.priority,
                operator: "notContains",
                text: cf.text,
                dxfId: id,
                formula: formula,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "beginsWith": {
              const firstCell = getFirstAddress(cf.sqref);
              const fomula = `LEFT(${firstCell},LEN("${cf.text}"))="${cf.text}"`;
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "beginsWith",
                sqref: cf.sqref,
                priority: cf.priority,
                operator: "beginsWith",
                text: cf.text,
                dxfId: id,
                formula: fomula,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "endsWith": {
              const firstCell = getFirstAddress(cf.sqref);
              const fomula = `RIGHT(${firstCell},LEN("${cf.text}"))="${cf.text}"`;
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "endsWith",
                sqref: cf.sqref,
                priority: cf.priority,
                operator: "endsWith",
                text: cf.text,
                dxfId: id,
                formula: fomula,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            case "timePeriod": {
              const firstCell = getFirstAddress(cf.sqref);
              let formula: string;

              switch (cf.timePeriod) {
                case "yesterday": {
                  formula = `FLOOR(${firstCell},1)=TODAY()-1`;
                  break;
                }
                case "today": {
                  formula = `FLOOR(${firstCell},1)=TODAY()`;
                  break;
                }
                case "tomorrow": {
                  formula = `FLOOR(${firstCell},1)=TODAY()+1`;
                  break;
                }
                case "last7Days": {
                  formula = `AND(TODAY()-FLOOR(${firstCell},1)&lt;=6,FLOOR(${firstCell},1)&lt;=TODAY())`;
                  break;
                }
                case "lastWeek": {
                  formula = `AND(TODAY()-ROUNDDOWN(${firstCell},0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${firstCell},0)&lt;(WEEKDAY(TODAY())+7))`;
                  break;
                }
                case "thisWeek": {
                  formula = `AND(TODAY()-ROUNDDOWN(${firstCell},0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(${firstCell},0)-TODAY()&lt;=7-WEEKDAY(TODAY()))`;
                  break;
                }
                case "nextWeek": {
                  formula = `AND(ROUNDDOWN(${firstCell},0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(${firstCell},0)-TODAY()&lt;(15-WEEKDAY(TODAY())))`;
                  break;
                }
                case "lastMonth": {
                  formula = `AND(MONTH(${firstCell})=MONTH(EDATE(TODAY(),0-1)),YEAR(${firstCell})=YEAR(EDATE(TODAY(),0-1)))`;
                  break;
                }
                case "thisMonth": {
                  formula = `AND(MONTH(${firstCell})=MONTH(TODAY()),YEAR(${firstCell})=YEAR(TODAY()))`;
                  break;
                }
                case "nextMonth": {
                  formula = `AND(MONTH(${firstCell})=MONTH(EDATE(TODAY(),0+1)),YEAR(${firstCell})=YEAR(EDATE(TODAY(),0+1)))`;
                  break;
                }
              }
              const conditionalFormatting: XlsxConditionalFormatting = {
                type: "timePeriod",
                sqref: cf.sqref,
                priority: cf.priority,
                timePeriod: cf.timePeriod,
                formula: formula,
                dxfId: id,
              };
              xcfs.push(conditionalFormatting);
              break;
            }
            default: {
              const _exhaustiveCheck: never = cf;
              throw new Error(
                `unknown conditional formatting type: ${_exhaustiveCheck}`
              );
            }
          }
        }
      }
      return xcfs;
    },

    makeXmlElm(formattings: XlsxConditionalFormatting[]) {
      let xml = "";

      for (const formatting of formattings) {
        switch (formatting.type) {
          case "top10": {
            const percent = formatting.percent ? ' percent="1"' : "";
            const bottom = formatting.bottom ? ' bottom="1"' : "";
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="top10" dxfId="${formatting.dxfId}" priority="${formatting.priority}"${percent}${bottom} rank="${formatting.rank}"/>` +
              "</conditionalFormatting>";
            break;
          }
          case "aboveAverage": {
            const aboveAverage = formatting.aboveAverage
              ? ""
              : ' aboveAverage="0"';
            const equalAverage = formatting.equalAverage
              ? ' equalAverage="1"'
              : "";
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="aboveAverage" dxfId="${formatting.dxfId}" priority="${formatting.priority}"${aboveAverage}${equalAverage}/>` +
              "</conditionalFormatting>";
            break;
          }
          case "duplicateValues": {
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="duplicateValues" dxfId="${formatting.dxfId}" priority="${formatting.priority}"/>` +
              "</conditionalFormatting>";
            break;
          }
          case "cellIs": {
            let formula: string;
            if (formatting.operator === "between") {
              formula = `<formula>${formatting.formulaA}</formula><formula>${formatting.formulaB}</formula>`;
            } else {
              formula = `<formula>${formatting.formula}</formula>`;
            }
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="cellIs" dxfId="${formatting.dxfId}" priority="${formatting.priority}" operator="${formatting.operator}">` +
              formula +
              `</cfRule>` +
              "</conditionalFormatting>";
            break;
          }
          case "containsText":
          case "notContainsText":
          case "beginsWith":
          case "endsWith": {
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="${formatting.type}" dxfId="${formatting.dxfId}" priority="${formatting.priority}" operator="${formatting.operator}" text="${formatting.text}">` +
              `<formula>${formatting.formula}</formula>` +
              `</cfRule>` +
              "</conditionalFormatting>";
            break;
          }
          case "timePeriod": {
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="timePeriod" dxfId="${formatting.dxfId}" priority="${formatting.priority}" timePeriod="${formatting.timePeriod}">` +
              `<formula>${formatting.formula}</formula>` +
              "</cfRule>" +
              "</conditionalFormatting>";
            break;
          }
          case "dataBar": {
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="dataBar" priority="${formatting.priority}">` +
              `<dataBar>` +
              `<cfvo type="min"/>` +
              `<cfvo type="max"/>` +
              `<color rgb="${formatting.color}"/>` +
              `</dataBar>` +
              `<extLst>` +
              `<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">` +
              `<x14:id>{${formatting.x14Id}}</x14:id>` +
              `</ext>` +
              `</extLst>` +
              `</cfRule>` +
              `</conditionalFormatting>`;
            break;
          }
          case "colorScale": {
            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="colorScale" priority="${formatting.priority}">` +
              `<colorScale>`;

            xml += '<cfvo type="min"/>';
            if ("mid" in formatting.colorScale) {
              xml += `<cfvo type="percentile" val="50"/>`;
            }
            xml += '<cfvo type="max"/>';

            for (const color of Object.values(formatting.colorScale)) {
              xml += `<color rgb="${color}"/>`;
            }

            xml += `</colorScale></cfRule></conditionalFormatting>`;
            break;
          }
          case "iconSet": {
            let iconSet;
            switch (formatting.iconSet) {
              case "3Arrows":
              case "3ArrowsGray":
              case "3Symbols":
              case "3Symbols2":
              case "3Flags": {
                iconSet =
                  `<iconSet iconSet="${formatting.iconSet}">` +
                  '<cfvo type="percent" val="0"/>' +
                  '<cfvo type="percent" val="33"/>' +
                  '<cfvo type="percent" val="67"/>' +
                  "</iconSet>";
                break;
              }
              case "4Arrows":
              case "4ArrowsGray": {
                iconSet =
                  `<iconSet iconSet="${formatting.iconSet}">` +
                  '<cfvo type="percent" val="0"/>' +
                  '<cfvo type="percent" val="25"/>' +
                  '<cfvo type="percent" val="50"/>' +
                  '<cfvo type="percent" val="75"/>' +
                  "</iconSet>";
                break;
              }
              case "5Arrows":
              case "5ArrowsGray": {
                iconSet =
                  `<iconSet iconSet="${formatting.iconSet}">` +
                  '<cfvo type="percent" val="0"/>' +
                  '<cfvo type="percent" val="20"/>' +
                  '<cfvo type="percent" val="40"/>' +
                  '<cfvo type="percent" val="60"/>' +
                  '<cfvo type="percent" val="80"/>' +
                  "</iconSet>";
                break;
              }
            }

            xml +=
              `<conditionalFormatting sqref="${formatting.sqref}">` +
              `<cfRule type="iconSet" priority="${formatting.priority}">` +
              iconSet +
              `</cfRule>` +
              `</conditionalFormatting>`;
            break;
          }
          default: {
            const _exhaustiveCheck: never = formatting;
            throw new Error(
              `unknown conditional formatting type: ${_exhaustiveCheck}`
            );
          }
        }
      }

      return xml;
    },
  };
}
