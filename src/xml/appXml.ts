export function makeAppXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    "<Application>xlsx-kaku</Application>" +
    "<Manager></Manager>" +
    "<Company></Company>" +
    "<HyperlinkBase></HyperlinkBase>" +
    "<AppVersion>16.0300</AppVersion>" +
    "</Properties>"
  );
}
