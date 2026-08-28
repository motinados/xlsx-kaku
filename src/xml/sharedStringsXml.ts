import { SharedStrings } from "../sharedStrings";
import { escapeXmlText } from "../utils";

export function makeSharedStringsXml(sharedStrings: SharedStrings) {
  if (sharedStrings.count === 0) {
    return null;
  }

  let result = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  result += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.count}" uniqueCount="${sharedStrings.uniqueCount}">`;
  for (const str of sharedStrings.getValuesInOrder()) {
    // Excel trims the leading and trailing whitespace unless it is preserved.
    const openingTag = /^\s|\s$/.test(str) ? '<t xml:space="preserve">' : "<t>";
    result += `<si>${openingTag}${escapeXmlText(str)}</t></si>`;
  }
  result += `</sst>`;
  return result;
}
