import { AutoFilter } from "../worksheet";
import {
  convColIndexToColName,
  convColNameToColIndex,
  devideAddress,
  escapeXmlText,
  quoteSheetName,
} from "../utils";

/**
 * The name Excel reserves for the range an auto filter is applied to.
 */
export const FILTER_DATABASE_NAME = "_xlnm._FilterDatabase";

const AUTO_FILTER_REF_PATTERN = /^[A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*$/;

/** The last column of a worksheet is "XFD". */
const MAX_COL_INDEX = 16383;
const MAX_ROW_NUMBER = 1048576;

export type AutoFilterModule = {
  name: string;
  getAutoFilter(): AutoFilter | null;
  set(autoFilter: AutoFilter): void;
  makeXmlElm(): string;
  makeDefinedNameElm(sheetName: string, localSheetId: number): string;
};

export function autoFilterModule(): AutoFilterModule {
  // A worksheet can have at most one auto filter.
  let autoFilter: AutoFilter | null = null;

  return {
    name: "auto-filter",
    getAutoFilter() {
      return autoFilter;
    },
    set(newAutoFilter: AutoFilter) {
      autoFilter = {
        ...newAutoFilter,
        ref: normalizeAutoFilterRef(newAutoFilter.ref),
      };
    },
    makeXmlElm() {
      if (autoFilter === null) {
        return "";
      }

      return `<autoFilter ref="${autoFilter.ref}"/>`;
    },
    makeDefinedNameElm(sheetName: string, localSheetId: number) {
      if (autoFilter === null) {
        return "";
      }

      const reference = `${quoteSheetName(sheetName)}!${convRefToAbsolute(
        autoFilter.ref
      )}`;

      return (
        `<definedName name="${FILTER_DATABASE_NAME}" localSheetId="${localSheetId}" hidden="1">` +
        escapeXmlText(reference) +
        "</definedName>"
      );
    },
  };
}

/**
 * Excel stores the range with the top-left address first, so the given range is
 * reordered to match it.
 *
 * e.g. "C10:A1" => "A1:C10"
 */
export function normalizeAutoFilterRef(ref: string): string {
  if (!AUTO_FILTER_REF_PATTERN.test(ref)) {
    throw new Error(
      `Invalid auto filter ref: "${ref}". It must be an uppercase range such as "A1:C10".`
    );
  }

  const [start, end] = ref.split(":") as [string, string];
  const [startColName, startRowNumber] = devideAddress(start);
  const [endColName, endRowNumber] = devideAddress(end);

  const startColIndex = convColNameToColIndex(startColName);
  const endColIndex = convColNameToColIndex(endColName);

  const maxColIndex = Math.max(startColIndex, endColIndex);
  const maxRowNumber = Math.max(startRowNumber, endRowNumber);
  if (maxColIndex > MAX_COL_INDEX || maxRowNumber > MAX_ROW_NUMBER) {
    throw new Error(
      `Auto filter ref "${ref}" is out of the range of a worksheet. The last cell is "XFD${MAX_ROW_NUMBER}".`
    );
  }

  const minColName = convColIndexToColName(
    Math.min(startColIndex, endColIndex)
  );
  const maxColName = convColIndexToColName(maxColIndex);
  const minRowNumber = Math.min(startRowNumber, endRowNumber);

  return `${minColName}${minRowNumber}:${maxColName}${maxRowNumber}`;
}

/**
 * e.g. "A1:C10" => "$A$1:$C$10"
 */
export function convRefToAbsolute(ref: string): string {
  return ref
    .split(":")
    .map((address) => {
      const [colName, rowNumber] = devideAddress(address);
      return `$${colName}$${rowNumber}`;
    })
    .join(":");
}
