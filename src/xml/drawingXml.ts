import { XlsxImage } from "./worksheetXml";

/**
 * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from>
            <xdr:col>0</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>0</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>2</xdr:col>
            <xdr:colOff>342900</xdr:colOff>
            <xdr:row>10</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:to>
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id="2" name="図 1">
                    <a:extLst>
                        <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{30326E85-BE84-6669-3F03-22BBDB7D5735}"/>
                        </a:ext>
                    </a:extLst>
                </xdr:cNvPr>
                <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                </xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="1714500" cy="1714500"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
            </xdr:spPr>
        </xdr:pic>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>
 */
export function makeDrawingXml(xlsxImages: XlsxImage[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml +=
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';

  for (const xlsxImage of xlsxImages) {
    xml += '<xdr:twoCellAnchor editAs="oneCell">';
    xml += "<xdr:from>";
    xml += `<xdr:col>${xlsxImage.from.col}</xdr:col>`;
    xml += `<xdr:colOff>${xlsxImage.from.colOff}</xdr:colOff>`;
    xml += `<xdr:row>${xlsxImage.from.row}</xdr:row>`;
    xml += `<xdr:rowOff>${xlsxImage.from.rowOff}</xdr:rowOff>`;
    xml += "</xdr:from>";
    //   xml += "<xdr:to>";
    //   xml += `<xdr:col>${xlsxImage.to.col}</xdr:col>`;
    //   xml += `<xdr:colOff>${xlsxImage.to.colOff}</xdr:colOff>`;
    //   xml += `<xdr:row>${xlsxImage.to.row}</xdr:row>`;
    //   xml += `<xdr:rowOff>${xlsxImage.to.rowOff}</xdr:rowOff>`;
    //   xml += "</xdr:to>";
    xml += "<xdr:pic>";
    xml += "<xdr:nvPicPr>";
    xml += '<xdr:cNvPr id="2" name="図 1">';
    xml += "<a:extLst>";
    xml += '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">';
    xml +=
      '<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{30326E85-BE84-6669-3F03-22BBDB7D5735}"/>';
    xml += "</a:ext>";
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
    xml += "</xdr:twoCellAnchor>";
  }

  xml += "</xdr:wsDr>";

  return xml;
}
