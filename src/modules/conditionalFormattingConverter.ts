import { Dxf } from "../dxf";
import { createUuid, getFirstAddress } from "../utils";
import { ConditionalFormatting } from "../worksheet";
import { XlsxConditionalFormatting } from "../xml/worksheetXml";

export function createXlsxConditionalFormattings(
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
          x14Id: createUuid(),
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
              cf.type === "atOrAboveAverage" || cf.type === "atOrBelowAverage",
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
}
