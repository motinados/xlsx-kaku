import { v4 as uuidv4 } from "uuid";
import { XlsxImage } from "./worksheetXml";

export function makeDrawingXml(xlsxImages: XlsxImage[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml +=
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';

  let prevId: string | null = null;
  for (const xlsxImage of xlsxImages) {
    const id = uuidv4();
    xml += '<xdr:oneCellAnchor editAs="oneCell">';
    xml += "<xdr:from>";
    xml += `<xdr:col>${xlsxImage.from.col}</xdr:col>`;
    xml += `<xdr:colOff>${xlsxImage.from.colOff}</xdr:colOff>`;
    xml += `<xdr:row>${xlsxImage.from.row}</xdr:row>`;
    xml += `<xdr:rowOff>${xlsxImage.from.rowOff}</xdr:rowOff>`;
    xml += "</xdr:from>";
    xml += `<xdr:ext cx="${xlsxImage.ext.cx}" cy="${xlsxImage.ext.cy}"/>`;
    xml += "<xdr:pic>";
    xml += "<xdr:nvPicPr>";
    xml += `<xdr:cNvPr id="${xlsxImage.id}" name="${xlsxImage.name}">`;
    xml += "<a:extLst>";
    xml += '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">';
    xml += `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{${id}}"/>`;
    xml += "</a:ext>";

    if (prevId) {
      xml += '<a:ext uri="{147F2762-F138-4A5C-976F-8EAC2B608ADB}">';
      xml += `<a16:predDERef xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" pred="{${prevId}}"/>`;
      xml += "</a:ext>";
    }

    xml += "</a:extLst>";
    xml += "</xdr:cNvPr>";
    xml += "<xdr:cNvPicPr>";
    xml += '<a:picLocks noChangeAspect="1"/>';
    xml += "</xdr:cNvPicPr>";
    xml += "</xdr:nvPicPr>";
    xml += "<xdr:blipFill>";
    xml += `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${xlsxImage.rId}"/>`;
    xml += "<a:stretch>";
    xml += "<a:fillRect/>";
    xml += "</a:stretch>";
    xml += "</xdr:blipFill>";
    xml += "<xdr:spPr>";
    xml += "<a:xfrm>";
    xml += '<a:off x="0" y="0"/>';
    xml += `<a:ext cx="${xlsxImage.ext.cx}" cy="${xlsxImage.ext.cy}"/>`;
    xml += "</a:xfrm>";
    xml += '<a:prstGeom prst="rect">';
    xml += "<a:avLst/>";
    xml += "</a:prstGeom>";
    xml += "</xdr:spPr>";
    xml += "</xdr:pic>";
    xml += "<xdr:clientData/>";
    xml += "</xdr:oneCellAnchor>";

    prevId = id;
  }

  xml += "</xdr:wsDr>";

  return xml;
}
