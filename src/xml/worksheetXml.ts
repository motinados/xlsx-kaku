import { v4 as uuidv4 } from "uuid";
import { FreezePane, MergeCell, Worksheet } from "..";
import { Cell, CellStyle, RowData, SheetData } from "../sheetData";
import { StyleMappers } from "../writer";
import {
  convColIndexToColName,
  convColNameToColIndex,
  getFirstAddress,
} from "../utils";
import { Alignment, CellXf } from "../cellXfs";
import { Hyperlinks } from "../hyperlinks";
import {
  ColProps,
  ConditionalFormatting,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  Image,
  RowProps,
} from "../worksheet";
import { Dxf } from "../dxf";
import { DrawingRels } from "../drawingRels";

export type XlsxCol = {
  /** e.g. column A is 0 */
  index: number;
  width: number;
  customWidth: boolean;
  cellXfId: number | null;
};

type XlsxRow = {
  /** e.g. rows[0] is 0 */
  index: number;
  height: number;
  customHeight: boolean;
  cellXfId: number | null;
};

type XlsxCellStyle = {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: Alignment;
};

type XlsxCell =
  | {
      type: "number";
      colName: string;
      rowNumber: number;
      value: number;
      cellXfId: number | null;
    }
  | {
      type: "string";
      colName: string;
      rowNumber: number;
      value: string;
      sharedStringId: number;
      cellXfId: number | null;
    }
  | {
      type: "date";
      colName: string;
      rowNumber: number;
      value: string;
      cellXfId: number | null;
    }
  | {
      type: "hyperlink";
      colName: string;
      rowNumber: number;
      value: string;
      sharedStringId: number;
      cellXfId: number | null;
    }
  | {
      type: "boolean";
      colName: string;
      rowNumber: number;
      value: boolean;
      cellXfId: number | null;
    }
  | {
      type: "formula";
      colName: string;
      rowNumber: number;
      value: string;
      cellXfId: number | null;
    }
  | {
      type: "merged";
      colName: string;
      rowNumber: number;
      cellXfId: number | null;
    };

export type GroupedXlsxCol = {
  startIndex: number;
  endIndex: number;
  width: number;
  customWidth: boolean;
  cellXfId: number | null;
};

export type XlsxConditionalFormatting =
  | {
      type: "top10";
      sqref: string;
      dxfId: number;
      priority: number;
      percent: boolean;
      bottom: boolean;
      rank: number;
    }
  | {
      type: "aboveAverage";
      sqref: string;
      dxfId: number;
      priority: number;
      aboveAverage: boolean;
      equalAverage: boolean;
    }
  | {
      type: "duplicateValues";
      sqref: string;
      dxfId: number;
      priority: number;
    }
  | {
      type: "cellIs";
      sqref: string;
      dxfId: number;
      priority: number;
      operator: "greaterThan" | "lessThan" | "equal";
      formula: string;
    }
  | {
      type: "cellIs";
      sqref: string;
      dxfId: number;
      priority: number;
      operator: "between";
      formulaA: string;
      formulaB: string;
    }
  | {
      type: "containsText";
      sqref: string;
      dxfId: number;
      priority: number;
      operator: "containsText";
      text: string;
      formula: string;
    }
  | {
      type: "notContainsText";
      sqref: string;
      dxfId: number;
      priority: number;
      operator: "notContains";
      text: string;
      formula: string;
    }
  | {
      type: "beginsWith";
      sqref: string;
      dxfId: number;
      priority: number;
      operator: "beginsWith";
      text: string;
      formula: string;
    }
  | {
      type: "endsWith";
      sqref: string;
      dxfId: number;
      priority: number;
      operator: "endsWith";
      text: string;
      formula: string;
    }
  | {
      type: "timePeriod";
      sqref: string;
      dxfId: number;
      priority: number;
      timePeriod:
        | "yesterday"
        | "today"
        | "tomorrow"
        | "last7Days"
        | "lastWeek"
        | "thisWeek"
        | "nextWeek"
        | "lastMonth"
        | "thisMonth"
        | "nextMonth";
      formula: string;
    }
  | {
      type: "dataBar";
      sqref: string;
      priority: number;
      color: string;
      x14Id: string;
      border: boolean;
      gradient: boolean;
      negativeBarBorderColorSameAsPositive: boolean;
    }
  | {
      type: "colorScale";
      sqref: string;
      priority: number;
      colorScale:
        | { min: string; max: string }
        | { min: string; mid: string; max: string };
    }
  | {
      type: "iconSet";
      sqref: string;
      priority: number;
      iconSet:
        | "3Arrows"
        | "4Arrows"
        | "5Arrows"
        | "3ArrowsGray"
        | "4ArrowsGray"
        | "5ArrowsGray"
        | "3Symbols"
        | "3Symbols2"
        | "3Flags";
    };

export type XlsxImage = {
  rId: string;
  id: string;
  name: string;
  editAs: "oneCell";
  from: {
    col: number;
    colOff: number;
    row: number;
    rowOff: number;
  };
  ext: {
    cx: number;
    cy: number;
  };
};

export function makeWorksheetXml(
  worksheet: Worksheet,
  styleMappers: StyleMappers,
  dxf: Dxf,
  drawingRels: DrawingRels,
  sheetCnt: number
) {
  styleMappers.hyperlinks.reset();
  styleMappers.worksheetRels.reset();

  const defaultColWidth = worksheet.props.defaultColWidth;
  const defaultRowHeight = worksheet.props.defaultRowHeight;
  const sheetData = worksheet.sheetData;

  const xlsxCols = new Map<number, XlsxCol>();
  for (const col of worksheet.cols.values()) {
    const xlsxCol = createXlsxColFromColProps(
      col,
      styleMappers,
      defaultColWidth
    );
    xlsxCols.set(xlsxCol.index, xlsxCol);
  }

  const xlsxRows = new Map<number, XlsxRow>();
  for (const row of worksheet.rows.values()) {
    const xlsxRow = createXlsxRowFromRowProps(row, styleMappers);
    xlsxRows.set(xlsxRow.index, xlsxRow);
  }

  const { spanStartNumber, spanEndNumber } = getSpansFromSheetData(sheetData);

  const colsElm = makeColsElm(groupXlsxCols(xlsxCols), defaultColWidth);
  const mergeCellsElm = makeMergeCellsElm(worksheet.mergeCells);

  const xlsxConditionalFormattings = createXlsxConditionalFormatting(
    worksheet.conditionalFormattings,
    dxf
  );
  const conditionalFormattingElm = makeConditionalFormattingElm(
    xlsxConditionalFormattings
  );

  const xlsxImages = worksheet.images.map((image) =>
    createXlsxImage(image, drawingRels)
  );

  let drawingRId: string | null = null;
  if (xlsxImages.length > 0) {
    drawingRId = styleMappers.worksheetRels.addWorksheetRel({
      target: `../drawings/drawing${sheetCnt + 1}.xml`,
      targetMode: null,
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    });
  }

  const sheetDataElm = makeSheetDataElm(
    sheetData,
    spanStartNumber,
    spanEndNumber,
    styleMappers,
    xlsxCols,
    xlsxRows
  );
  const dimension = getDimension(sheetData, spanStartNumber, spanEndNumber);
  const tabSelected = sheetCnt === 0;
  const sheetViewsElm = makeSheetViewsElm(
    tabSelected,
    dimension,
    worksheet.freezePane
  );
  const sheetFormatPrElm = makeSheetFormatPrElm(
    defaultRowHeight,
    defaultColWidth
  );
  const drawingElm = makeDrawingElm(drawingRId);
  const extLstElm = makeExtLstElm(xlsxConditionalFormattings);

  // Perhaps passing a UUID to every sheet won't cause any issues,
  // but for the sake of integration testing, only the first sheet is given specific UUID.
  const uuid =
    sheetCnt === 0 ? "00000000-0001-0000-0000-000000000000" : uuidv4();
  const sheetXml = composeSheetXml(
    uuid,
    colsElm,
    sheetViewsElm,
    sheetFormatPrElm,
    sheetDataElm,
    mergeCellsElm,
    conditionalFormattingElm,
    extLstElm,
    drawingElm,
    dimension,
    styleMappers.hyperlinks
  );

  let worksheetRels;
  if (styleMappers.worksheetRels.relsLength > 0) {
    worksheetRels = styleMappers.worksheetRels.makeXML();
  } else {
    worksheetRels = null;
  }

  let drawingRelsXml;
  if (drawingRels.rels.length > 0) {
    drawingRelsXml = drawingRels.makeXml();
  } else {
    drawingRelsXml = null;
  }

  return {
    sheetXml,
    worksheetRels,
    drawingRelsXml,
    xlsxImages,
  };
}

export function createXlsxColFromColProps(
  col: ColProps,
  mappers: StyleMappers,
  defaultWidth: number
): XlsxCol {
  let cellXfId: number | null = null;
  if (col.style) {
    const style = composeXlsxCellStyle(col.style, mappers);
    if (style === null) {
      throw new Error("style is null");
    }
    cellXfId = mappers.cellXfs.getCellXfId(style);
  }

  return {
    index: col.index,
    width: col.width ?? defaultWidth,
    customWidth: col.width !== undefined && col.width !== defaultWidth,
    cellXfId: cellXfId,
  };
}

export function composeXlsxCellStyle(
  style: CellStyle | undefined,
  mappers: StyleMappers
): XlsxCellStyle | null {
  if (style) {
    const _style: XlsxCellStyle = {
      fillId: style.fill ? mappers.fills.getFillId(style.fill) : 0,
      fontId: style.font ? mappers.fonts.getFontId(style.font) : 0,
      borderId: style.border ? mappers.borders.getBorderId(style.border) : 0,
      numFmtId: style.numberFormat
        ? mappers.numberFormats.getNumFmtId(style.numberFormat.formatCode)
        : 0,
    };

    if (style.alignment) {
      _style.alignment = style.alignment;
    }

    return _style;
  }

  return null;
}

export function createXlsxRowFromRowProps(
  row: RowProps,
  styleMappers: StyleMappers
): XlsxRow {
  let cellXfId: number | null = null;
  if (row.style) {
    const style = composeXlsxCellStyle(row.style, styleMappers);
    if (style === null) {
      throw new Error("style is null");
    }
    cellXfId = styleMappers.cellXfs.getCellXfId(style);
  }

  return {
    index: row.index,
    height: row.height ?? DEFAULT_ROW_HEIGHT,
    customHeight: row.height !== undefined && row.height !== DEFAULT_ROW_HEIGHT,
    cellXfId: cellXfId,
  };
}

function createXlsxConditionalFormatting(
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

export function createXlsxImage(
  image: Image,
  drawingRels: DrawingRels
): XlsxImage {
  const num = drawingRels.length + 1;
  const rId = drawingRels.addDrawingRel({
    target: `../media/image${num}.png`,
    relationshipType:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  });

  return {
    rId,
    // Files created in online Excel seem to start with a sequential number begginning from 2.
    id: String(num + 1),
    name: image.displayName,
    editAs: "oneCell",
    from: {
      col: 0,
      colOff: 0,
      row: 0,
      rowOff: 0,
    },
    ext: {
      cx: (914400 / 96) * image.width,
      cy: (914400 / 96) * image.height,
    },
  };
}

export function isEqualsXlsxCol(a: XlsxCol, b: XlsxCol) {
  return (
    // index must not be compared.
    a.width === b.width &&
    a.customWidth === b.customWidth &&
    a.cellXfId === b.cellXfId
  );
}

export function groupXlsxCols(cols: Map<number, XlsxCol>) {
  const result: GroupedXlsxCol[] = [];
  let startCol: XlsxCol;
  let endCol: XlsxCol;

  let i = 0;
  for (const col of cols.values()) {
    if (i === 0) {
      // the first
      startCol = col;
      endCol = col;
    } else {
      if (isEqualsXlsxCol(startCol!, col)) {
        endCol = col;
      } else {
        result.push({
          startIndex: startCol!.index,
          endIndex: endCol!.index,
          width: startCol!.width,
          customWidth: startCol!.customWidth,
          cellXfId: startCol!.cellXfId,
        });
        startCol = col;
        endCol = col;
      }
    }

    if (i == cols.size - 1) {
      // the last
      result.push({
        startIndex: startCol!.index,
        endIndex: endCol!.index,
        width: startCol!.width,
        customWidth: startCol!.customWidth,
        cellXfId: startCol!.cellXfId,
      });
    }

    i++;
  }

  return result;
}

export function makeColsElm(
  cols: GroupedXlsxCol[],
  defaultColWidth: number
): string {
  if (cols.length === 0) {
    return "";
  }

  let result = "<cols>";
  for (const col of cols) {
    result += `<col min="${col.startIndex + 1}" max="${col.endIndex + 1}"`;

    if (col.customWidth) {
      result += ` width="${col.width}" customWidth="1"`;
    } else {
      result += ` width="${defaultColWidth}"`;
    }

    if (col.cellXfId) {
      result += ` style="${col.cellXfId}"`;
    }

    result += "/>";
  }
  result += "</cols>";

  return result;
}

export function makeMergeCellsElm(mergeCells: MergeCell[]) {
  if (mergeCells.length === 0) {
    return "";
  }

  let result = `<mergeCells count="${mergeCells.length}">`;
  for (const mergeCell of mergeCells) {
    result += `<mergeCell ref="${mergeCell.ref}"/>`;
  }
  result += "</mergeCells>";

  return result;
}

export function makeConditionalFormattingElm(
  formattings: XlsxConditionalFormatting[]
): string {
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
        const aboveAverage = formatting.aboveAverage ? "" : ' aboveAverage="0"';
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
}

export function makeSheetDataElm(
  sheetData: SheetData,
  spanStartNumber: number,
  spanEndNumber: number,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  xlsxRows: Map<number, XlsxRow>
) {
  let result = `<sheetData>`;
  let rowIndex = 0;
  for (const row of sheetData) {
    const str = makeRowElm(
      row,
      rowIndex,
      spanStartNumber,
      spanEndNumber,
      styleMappers,
      xlsxCols,
      xlsxRows
    );
    result += str;
    rowIndex++;
  }
  result += `</sheetData>`;
  return result;
}

export function getSpansFromSheetData(sheetData: SheetData) {
  const all = sheetData
    .map((row) => {
      return getSpans(row);
    })
    .filter((row) => row !== null) as {
    startNumber: number;
    endNumber: number;
  }[];
  const minStartNumber = Math.min(...all.map((row) => row.startNumber));
  const maxEndNumber = Math.max(...all.map((row) => row.endNumber));
  return { spanStartNumber: minStartNumber, spanEndNumber: maxEndNumber };
}

export function getSpans(row: RowData) {
  const first = findFirstNonNullCell(row);
  if (first === undefined || first === null) {
    return null;
  }

  const last = findLastNonNullCell(row);
  if (last === undefined) {
    return null;
  }

  const startNumber = first.index + 1;
  const endNumber = last.index + 1;

  return { startNumber, endNumber };
}

export function findFirstNonNullCell(row: RowData) {
  let index = 0;
  let firstNonNullCell = null;
  for (let i = 0; i < row.length; i++) {
    if (row[i] !== null) {
      firstNonNullCell = row[i]!;
      index = i;
      break;
    }
  }
  return { firstNonNullCell, index };
}

/**
 *  [null, null, null, nonnull, null] => index is 3
 */
export function findLastNonNullCell(row: RowData) {
  let index = 0;
  let lastNonNullCell = null;
  for (let i = row.length - 1; i >= 0; i--) {
    if (row[i] !== null) {
      lastNonNullCell = row[i]!;
      index = i;
      break;
    }
  }
  return { lastNonNullCell, index };
}

/**
 * <row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
 */
export function makeRowElm(
  row: RowData,
  rowIndex: number,
  spanStartNumber: number,
  spanEndNumber: number,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  xlsxRows: Map<number, XlsxRow>
): string {
  if (row.length === 0) {
    return "";
  }

  const rowNumber = rowIndex + 1;

  let result = `<row r="${rowNumber}" spans="${spanStartNumber}:${spanEndNumber}"`;

  const xlsxRow = xlsxRows.get(rowIndex);
  if (xlsxRow) {
    if (xlsxRow.cellXfId) {
      result += ` s="${xlsxRow.cellXfId}" customFormat="1"`;
    }

    if (xlsxRow.height && xlsxRow.customHeight) {
      result += ` ht="${xlsxRow.height}" customHeight="1"`;
    }
  }
  result += ">";

  let columnIndex = 0;
  for (const cell of row) {
    if (cell !== null) {
      result += makeCellElm(
        convertCellToXlsxCell(
          cell,
          columnIndex,
          rowIndex,
          styleMappers,
          xlsxCols,
          xlsxRow
        )
      );
    }

    columnIndex++;
  }

  result += `</row>`;
  return result;
}

export function makeCellElm(cell: XlsxCell) {
  switch (cell.type) {
    case "number": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s}><v>${cell.value}</v></c>`;
    }
    case "string": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s} t="s"><v>${cell.sharedStringId}</v></c>`;
    }
    case "date": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      const serialValue = convertIsoStringToSerialValue(cell.value);
      return `<c r="${cell.colName}${cell.rowNumber}"${s}><v>${serialValue}</v></c>`;
    }
    case "hyperlink": {
      const s = ` s="${cell.cellXfId}"`;
      return `<c r="${cell.colName}${cell.rowNumber}"${s} t="s"><v>${cell.sharedStringId}</v></c>`;
    }
    case "boolean": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      const v = cell.value ? 1 : 0;
      return `<c r="${cell.colName}${cell.rowNumber}"${s} t="b"><v>${v}</v></c>`;
    }
    case "formula": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s}><f>${cell.value}</f></c>`;
    }
    case "merged": {
      const s = cell.cellXfId ? ` s="${cell.cellXfId}"` : "";
      return `<c r="${cell.colName}${cell.rowNumber}"${s}/>`;
    }
    default: {
      const _exhaustiveCheck: never = cell;
      throw new Error(`unknown cell type: ${_exhaustiveCheck}`);
    }
  }
}

/**
 * https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
 * @param isoString
 * @returns
 */
export function convertIsoStringToSerialValue(isoString: string): number {
  const baseDate = new Date("1899-12-31T00:00:00.000Z");
  const targetDate = new Date(isoString);
  const differenceInDays =
    (targetDate.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24);
  // Excel uses January 0, 1900 as a base (which is actually December 31, 1899), so add 1 to the result
  return differenceInDays + 1;
}

export function convertCellToXlsxCell(
  cell: Cell,
  columnIndex: number,
  rowIndex: number,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  xlsxRow: XlsxRow | undefined
): XlsxCell {
  const rowNumber = rowIndex + 1;
  const colName = convColIndexToColName(columnIndex);

  switch (cell.type) {
    case "number": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
      );
      return {
        type: "number",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "string": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
      );
      const sharedStringId = styleMappers.sharedStrings.getIndex(cell.value);
      return {
        type: "string",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        sharedStringId: sharedStringId,
        cellXfId: cellXfId,
      };
    }
    case "date": {
      assignDateStyleIfUndefined(cell);
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
      );
      return {
        type: "date",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "hyperlink": {
      assignHyperlinkStyleIfUndefined(cell);
      const composedStyle = composeXlsxCellStyle(cell.style, styleMappers);
      if (composedStyle === null) {
        throw new Error("composedStyle is null for hyperlink");
      }
      const xfId = styleMappers.cellStyleXfs.getCellStyleXfId(composedStyle);
      if (xfId === null) {
        throw new Error("xfId is null for hyperlink");
      }

      const cellXf: CellXf = {
        xfId: xfId,
        ...composedStyle,
      };
      const cellXfId = styleMappers.cellXfs.getCellXfId(cellXf);
      const sharedStringId = styleMappers.sharedStrings.getIndex(cell.text);

      styleMappers.cellStyles.getCellStyleId({
        name: "Hyperlink",
        xfId: xfId,
        uid: "{00000000-000B-0000-0000-000008000000}",
      });

      if (cell.linkType === "external") {
        const rid = styleMappers.worksheetRels.addWorksheetRel({
          target: cell.value,
          targetMode: "External",
          relationshipType:
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        });
        styleMappers.hyperlinks.addHyperlink({
          linkType: "external",
          ref: `${colName}${rowNumber}`,
          rid: rid,
          uuid: uuidv4(),
        });
      } else if (cell.linkType === "internal") {
        styleMappers.hyperlinks.addHyperlink({
          linkType: "internal",
          ref: `${colName}${rowNumber}`,
          location: cell.value,
          display: cell.text,
          uuid: uuidv4(),
        });
      } else if (cell.linkType === "email") {
        const rid = styleMappers.worksheetRels.addWorksheetRel({
          target: `mailto:${cell.value}`,
          targetMode: "External",
          relationshipType:
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        });
        styleMappers.hyperlinks.addHyperlink({
          linkType: "email",
          ref: `${colName}${rowNumber}`,
          rid: rid,
          uuid: uuidv4(),
        });
      }

      return {
        type: "hyperlink",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        sharedStringId: sharedStringId,
        cellXfId: cellXfId,
      };
    }
    case "boolean": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
      );
      return {
        type: "boolean",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "formula": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
      );
      return {
        type: "formula",
        colName: colName,
        rowNumber: rowNumber,
        value: cell.value,
        cellXfId: cellXfId,
      };
    }
    case "merged": {
      const cellXfId = getCellXfId(
        cell,
        colName,
        styleMappers,
        xlsxCols,
        xlsxRow
      );
      return {
        type: "merged",
        colName: colName,
        rowNumber: rowNumber,
        cellXfId: cellXfId,
      };
    }
    default: {
      const _exhaustiveCheck: never = cell;
      throw new Error(`unknown cell type: ${_exhaustiveCheck}`);
    }
  }
}

function getCellXfId(
  cell: Cell,
  colName: string,
  styleMappers: StyleMappers,
  xlsxCols: Map<number, XlsxCol>,
  foundRow: XlsxRow | undefined
) {
  const composedStyle = composeXlsxCellStyle(cell.style, styleMappers);
  if (composedStyle) {
    return styleMappers.cellXfs.getCellXfId(composedStyle);
  }

  const foundCol = xlsxCols.get(convColNameToColIndex(colName));
  if (foundCol) {
    return foundCol.cellXfId;
  }

  if (foundRow) {
    return foundRow.cellXfId;
  }

  return null;
}

function assignDateStyleIfUndefined(cell: Cell) {
  if (cell.type === "date" && cell.style === undefined) {
    cell.style = { numberFormat: { formatCode: "yyyy-mm-dd" } };
  }
}

function assignHyperlinkStyleIfUndefined(cell: Cell) {
  if (cell.type === "hyperlink" && cell.style === undefined) {
    cell.style = {
      font: {
        name: "Calibri",
        size: 11,
        color: "0563c1",
        underline: "single",
      },
    };
  }
}

export function getDimension(
  sheetData: SheetData,
  spanStartNumber: number,
  spanEndNumber: number
) {
  // FIXME: The dimension is alse affected by 'cols'. It can have the correct value even without sheetData.
  if (sheetData.length === 0) {
    // This is a workaround for the case where sheetData is empty.
    return { start: "A1", end: "A1" };
  }

  const firstRowIndex = findFirstNotBlankRow(sheetData);
  const lastRowIndex = findLastNotBlankRow(sheetData);
  if (firstRowIndex === null || lastRowIndex === null) {
    throw new Error("sheetData is empty");
  }

  const firstRowNumber = firstRowIndex + 1;
  const lastRowNumber = lastRowIndex + 1;

  const firstColumn = convColIndexToColName(spanStartNumber - 1);
  const lastColumn = convColIndexToColName(spanEndNumber - 1);

  return {
    start: `${firstColumn}${firstRowNumber}`,
    end: `${lastColumn}${lastRowNumber}`,
  };
}

function findFirstNotBlankRow(sheetData: SheetData) {
  let index = 0;
  for (let i = 0; i < sheetData.length; i++) {
    const row = sheetData[i]!;
    if (row.length > 0) {
      index = i;
      break;
    }
  }
  return index;
}

function findLastNotBlankRow(sheetData: SheetData) {
  let index = 0;
  for (let i = sheetData.length - 1; i >= 0; i--) {
    const row = sheetData[i]!;
    if (row.length > 0) {
      index = i;
      break;
    }
  }
  return index;
}

// <sheetViews>
// <sheetView tabSelected="1" workbookViewId="0">
//     <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
//     <selection pane="topRight"/>
// </sheetView>
// </sheetViews>
export function makeSheetViewsElm(
  tabSelected: boolean,
  dimension: { start: string; end: string },
  freezePane: FreezePane | null
) {
  const openingTabSelectedTag = tabSelected
    ? `<sheetView tabSelected="1" workbookViewId="0">`
    : `<sheetView workbookViewId="0">`;

  if (freezePane === null) {
    let result =
      "<sheetViews>" +
      openingTabSelectedTag +
      `<selection activeCell="${dimension.start}" sqref="${dimension.start}"/>` +
      "</sheetView>" +
      "</sheetViews>";
    return result;
  }

  switch (freezePane.target) {
    case "column": {
      let result =
        "<sheetViews>" +
        openingTabSelectedTag +
        `<pane ySplit="${freezePane.split}" topLeftCell="A${
          freezePane.split + 1
        }" activePane="bottomLeft" state="frozen"/>` +
        `<selection pane="bottomLeft" activeCell="${dimension.start}" sqref="${dimension.start}"/>` +
        "</sheetView>" +
        "</sheetViews>";
      return result;
    }
    case "row": {
      let result =
        "<sheetViews>" +
        openingTabSelectedTag +
        `<pane xSplit="${
          freezePane.split
        }" topLeftCell="${convColIndexToColName(
          freezePane.split
        )}1" activePane="topRight" state="frozen"/>` +
        `<selection pane="topRight" activeCell="${dimension.start}" sqref="${dimension.start}"/>` +
        "</sheetView>" +
        "</sheetViews>";
      return result;
    }
    default: {
      const _exhaustiveCheck: never = freezePane.target;
      throw new Error(`unknown freezePane type: ${_exhaustiveCheck}`);
    }
  }
}

export function makeSheetFormatPrElm(
  defaultRowHeight: number,
  defaultColWidth: number
) {
  // There should be no issue with always the defaultColWidth,
  // but due to differences in integration tests with files created in Online Excel,
  // we deliberately avoid adding it when it's the same value as DEFAULT_COL_WIDTH.
  const shhetFormatPrXML =
    defaultColWidth === DEFAULT_COL_WIDTH
      ? `<sheetFormatPr defaultRowHeight="${defaultRowHeight}"/>`
      : `<sheetFormatPr defaultRowHeight="${defaultRowHeight}" defaultColWidth="${defaultColWidth}"/>`;

  return shhetFormatPrXML;
}

export function makeExtLstElm(
  xlsxConditionalFormattings: XlsxConditionalFormatting[]
) {
  const dataBars = xlsxConditionalFormattings.filter(
    (cf) => cf.type === "dataBar"
  );

  if (dataBars.length === 0) {
    return "";
  }

  let xml =
    "<extLst>" +
    '<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">' +
    "<x14:conditionalFormattings>";

  for (const formatting of dataBars) {
    if (formatting.type === "dataBar") {
      xml +=
        `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
        `<x14:cfRule type="dataBar" id="{${formatting.x14Id}}">` +
        `<x14:dataBar minLength="0" maxLength="100"${
          formatting.border ? ' border="1"' : ""
        }${formatting.gradient ? "" : ' gradient="0"'}${
          formatting.negativeBarBorderColorSameAsPositive
            ? ""
            : ' negativeBarBorderColorSameAsPositive="0"'
        }>` +
        `<x14:cfvo type="autoMin"/>` +
        `<x14:cfvo type="autoMax"/>` +
        `${
          formatting.border
            ? `<x14:borderColor rgb="${formatting.color}"/>`
            : ""
        }` +
        `<x14:negativeFillColor rgb="FFFF0000"/>` +
        `${
          formatting.border ? '<x14:negativeBorderColor rgb="FFFF0000"/>' : ""
        }` +
        `<x14:axisColor rgb="FF000000"/>` +
        `</x14:dataBar>` +
        `</x14:cfRule>` +
        `<xm:sqref>${formatting.sqref}</xm:sqref>` +
        `</x14:conditionalFormatting>`;
    }
  }

  xml += "</x14:conditionalFormattings></ext></extLst>";
  return xml;
}

export function makeDrawingElm(drawingRID: string | null) {
  if (drawingRID === null) {
    return "";
  }

  return `<drawing r:id="${drawingRID}"/>`;
}

export function composeSheetXml(
  uuid: string,
  colsElm: string,
  sheetViewsElm: string,
  sheetFormatPrElm: string,
  sheetDataString: string,
  mergeCellsElm: string,
  conditionalFormattingElm: string,
  extLstElm: string,
  drawingElm: string,
  dimension: { start: string; end: string },
  hyperlinks: Hyperlinks
) {
  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{${uuid}}">` +
    `<dimension ref="${dimension.start}:${dimension.end}"/>` +
    sheetViewsElm +
    sheetFormatPrElm +
    colsElm +
    sheetDataString;

  if (hyperlinks.getHyperlinks().length > 0) {
    result += hyperlinks.makeXML();
  }

  result +=
    mergeCellsElm +
    conditionalFormattingElm +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    drawingElm +
    extLstElm +
    "</worksheet>";

  return result;
}
