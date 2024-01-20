export function makeCoreXml() {
  const isoDate = new Date().toISOString();

  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:title></dc:title>" +
    "<dc:subject></dc:subject>" +
    "<dc:creator></dc:creator>" +
    "<cp:keywords></cp:keywords>" +
    "<dc:description></dc:description>" +
    "<cp:lastModifiedBy></cp:lastModifiedBy>" +
    "<cp:revision></cp:revision>" +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:modified><cp:category></cp:category>` +
    "<cp:contentStatus></cp:contentStatus>" +
    "</cp:coreProperties>";

  return result;
}
