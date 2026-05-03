import type { Borders } from "./borders";
import type { CellStyles } from "./cellStyles";
import type { CellStyleXfs } from "./cellStyleXfs";
import type { CellXfs } from "./cellXfs";
import type { Fills } from "./fills";
import type { Fonts } from "./fonts";
import type { Hyperlinks } from "./hyperlinks";
import type { NumberFormats } from "./numberFormats";
import type { SharedStrings } from "./sharedStrings";
import type { WorksheetRels } from "./worksheetRels";

export type WorkbookStyleMappers = {
  fills: Fills;
  fonts: Fonts;
  borders: Borders;
  numberFormats: NumberFormats;
  sharedStrings: SharedStrings;
  cellStyleXfs: CellStyleXfs;
  cellXfs: CellXfs;
  cellStyles: CellStyles;
};

export type WorksheetStyleMappers = {
  hyperlinks: Hyperlinks;
  worksheetRels: WorksheetRels;
};

export type StyleMappers = WorkbookStyleMappers & WorksheetStyleMappers;
