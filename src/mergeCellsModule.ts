import { MergeCell } from "./worksheet";
import { expandRange } from "./utils";
import { WorksheetType } from "./worksheet";

export type MergeCellsModule = {
  name: string;
  getMergeCells(): MergeCell[];
  add(worksheet: WorksheetType, mergeCell: MergeCell): void;
  makeXmlElm(): string;
};

export function mergeCellsModule(): MergeCellsModule {
  const mergeCells: MergeCell[] = [];
  return {
    name: "merge-cells",
    getMergeCells() {
      return mergeCells;
    },
    add(worksheet: WorksheetType, mergeCell: MergeCell) {
      // Within the range to be merged, cells are set with the type of "merged".
      const addresses = expandRange(mergeCell.ref);
      for (let i = 0; i < addresses.length; i++) {
        const address = addresses[i];
        if (address) {
          const [colIndex, rowIndex] = address;

          if (i === 0) {
            // If the cell is not set, set it as empty string.
            const cell = worksheet.getCell(rowIndex, colIndex);
            if (!cell) {
              worksheet.setCell(rowIndex, colIndex, {
                type: "string",
                value: "",
              });
            }
          } else {
            worksheet.setCell(rowIndex, colIndex, { type: "merged" });
          }
        }
      }
      mergeCells.push(mergeCell);
    },
    makeXmlElm() {
      if (mergeCells.length === 0) {
        return "";
      }

      let result = `<mergeCells count="${mergeCells.length}">`;
      for (const mergeCell of mergeCells) {
        result += `<mergeCell ref="${mergeCell.ref}"/>`;
      }
      result += "</mergeCells>";

      return result;
    },
  };
}
