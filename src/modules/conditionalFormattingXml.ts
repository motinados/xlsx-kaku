import { XlsxConditionalFormatting } from "../xml/worksheetXml";
import { escapeXmlAttribute, escapeXmlText } from "../utils";

export function makeConditionalFormattingXml(
  formattings: XlsxConditionalFormatting[]
) {
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
        const equalAverage = formatting.equalAverage ? ' equalAverage="1"' : "";
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
          formula =
            `<formula>${escapeXmlText(formatting.formulaA)}</formula>` +
            `<formula>${escapeXmlText(formatting.formulaB)}</formula>`;
        } else {
          formula = `<formula>${escapeXmlText(formatting.formula)}</formula>`;
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
          `<cfRule type="${formatting.type}" dxfId="${formatting.dxfId}" priority="${formatting.priority}" operator="${formatting.operator}" text="${escapeXmlAttribute(
            formatting.text
          )}">` +
          `<formula>${escapeXmlText(formatting.formula)}</formula>` +
          `</cfRule>` +
          "</conditionalFormatting>";
        break;
      }
      case "timePeriod": {
        xml +=
          `<conditionalFormatting sqref="${formatting.sqref}">` +
          `<cfRule type="timePeriod" dxfId="${formatting.dxfId}" priority="${formatting.priority}" timePeriod="${formatting.timePeriod}">` +
          `<formula>${escapeXmlText(formatting.formula)}</formula>` +
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
}
