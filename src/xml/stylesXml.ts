import { StyleMappers } from "../writer";

export function makeStylesXml(styleMappers: StyleMappers) {
  let result =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
    styleMappers.numberFormats.makeXml() +
    // results.push('<fonts count="1">');
    // results.push("<font>");
    // results.push('<sz val="11"/>');
    // results.push('<color theme="1"/>');
    // results.push('<name val="Calibri"/>');
    // results.push('<family val="2"/>');
    // results.push('<scheme val="minor"/></font>');
    // results.push("</fonts>");
    styleMappers.fonts.makeXml() +
    // results.push('<fills count="2">');
    // results.push('<fill><patternFill patternType="none"/></fill>');
    // results.push('<fill><patternFill patternType="gray125"/></fill>');
    // results.push("</fills>");
    styleMappers.fills.makeXml() +
    // results.push('<borders count="1">');
    // results.push("<border><left/><right/><top/><bottom/><diagonal/></border>");
    // results.push("</borders>");
    styleMappers.borders.makeXml() +
    // results.push(
    //   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    // );
    styleMappers.cellStyleXfs.makeXml() +
    // results.push(
    //   '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    // );
    styleMappers.cellXfs.makeXml() +
    // results.push(
    //   '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>'
    // );
    styleMappers.cellStyles.makeXml() +
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>' +
    "<extLst>" +
    '<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>' +
    '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
    '<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext>' +
    "</extLst>" +
    "</styleSheet>";

  return result;
}
