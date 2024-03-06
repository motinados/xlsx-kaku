import { v4 as uuidv4 } from "uuid";
import { Worksheet } from "../worksheet";

export function makeWorkbookXml(worksheets: Worksheet[]) {
  const documentId = uuidv4();

  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">' +
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="27123"/>' +
    '<workbookPr defaultThemeVersion="166925"/>' +
    `<xr:revisionPtr revIDLastSave="0" documentId="8_{${documentId}}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>` +
    "<bookViews>" +
    '<workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/>' +
    "</bookViews>" +
    "<sheets>";

  let sheetId = 1;
  for (const sheet of worksheets) {
    result += `<sheet name="${sheet.name}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`;
    sheetId++;
  }

  result +=
    "</sheets>" +
    "<extLst>" +
    '<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
    '<x15:workbookPr chartTrackingRefBase="1"/>' +
    "</ext>" +
    '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">' +
    "<xcalcf:calcFeatures>" +
    '<xcalcf:feature name="microsoft.com:RD"/>' +
    '<xcalcf:feature name="microsoft.com:Single"/>' +
    '<xcalcf:feature name="microsoft.com:FV"/>' +
    '<xcalcf:feature name="microsoft.com:CNMTM"/>' +
    '<xcalcf:feature name="microsoft.com:LET_WF"/>' +
    '<xcalcf:feature name="microsoft.com:LAMBDA_WF"/>' +
    '<xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/>' +
    "</xcalcf:calcFeatures>" +
    "</ext>" +
    "</extLst>" +
    "</workbook>";

  return result;
}
