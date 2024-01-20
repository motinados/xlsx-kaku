import { SharedStrings } from "../sharedStrings";

export function makeSharedStringsXml(sharedStrings: SharedStrings) {
  if (sharedStrings.count === 0) {
    return null;
  }

  let result = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  result += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.count}" uniqueCount="${sharedStrings.uniqueCount}">`;
  for (const str of sharedStrings.getValuesInOrder()) {
    result += `<si><t>${str}</t></si>`;
  }
  result += `</sst>`;
  return result;
}
