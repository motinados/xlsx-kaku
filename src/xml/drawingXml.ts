export function makeDrawingXml(drawingImageElm: string): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml +=
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
  xml += drawingImageElm;
  xml += "</xdr:wsDr>";

  return xml;
}
