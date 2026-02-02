import {
  ConditionalFormattingModule,
  conditionalFormattingModule,
} from "./modules/conditionalFormattingModule";
import { DxfStyle } from "./dxf";
import { ImageModule, imageModule } from "./modules/imageModule";
import { ImageStore } from "./imageStore";
import { MergeCellsModule, mergeCellsModule } from "./modules/mergeCellsModule";
import { CellStyle, NullableCell, SettableCell, SheetData } from "./sheetData";

/**
 * The value is the same as the one in files created with Online Excel.
 * Changing this value will result in differences in integration tests.
 */
export const DEFAULT_COL_WIDTH = 9;
export const DEFAULT_ROW_HEIGHT = 13.5;

export type ColOpts = {
  index: number;
  width?: number;
  style?: CellStyle;
};

export type RowOpts = {
  index: number;
  height?: number;
  style?: CellStyle;
};

export type MergeCell = {
  /**
   * e.g. "A2:B4"
   */
  ref: string;
};

export type FreezePane = {
  target: "column" | "row";
  split: number;
};

export type ConditionalFormatting =
  | {
      type: "top" | "bottom";
      sqref: string;
      priority: number;
      percent: boolean;
      rank: number;
      style: DxfStyle;
    }
  | {
      type:
        | "aboveAverage"
        | "belowAverage"
        | "atOrAboveAverage"
        | "atOrBelowAverage";
      sqref: string;
      priority: number;
      style: DxfStyle;
    }
  | {
      type: "duplicateValues";
      sqref: string;
      priority: number;
      style: DxfStyle;
    }
  | {
      type: "greaterThan" | "lessThan" | "equal";
      sqref: string;
      priority: number;
      formula: string | number;
      style: DxfStyle;
    }
  | {
      type: "between";
      sqref: string;
      priority: number;
      formulaA: string | number;
      formulaB: string | number;
      style: DxfStyle;
    }
  | {
      type: "containsText" | "notContainsText" | "beginsWith" | "endsWith";
      sqref: string;
      priority: number;
      text: string;
      style: DxfStyle;
    }
  | {
      type: "timePeriod";
      sqref: string;
      priority: number;
      timePeriod:
        | "yesterday"
        | "today"
        | "tomorrow"
        | "last7Days"
        | "thisMonth"
        | "lastMonth"
        | "nextMonth"
        | "thisWeek"
        | "lastWeek"
        | "nextWeek";
      style: DxfStyle;
    }
  | {
      type: "dataBar";
      sqref: string;
      priority: number;
      color: string;
      border: boolean;
      gradient: boolean;
      negativeBarBorderColorSameAsPositive: boolean;
    }
  | {
      type: "colorScale";
      sqref: string;
      priority: number;
      colorScale:
        | {
            min: string;
            max: string;
          }
        | {
            min: string;
            mid: string;
            max: string;
          };
    }
  | {
      type: "iconSet";
      sqref: string;
      priority: number;
      iconSet:
        | "3Arrows"
        | "4Arrows"
        | "5Arrows"
        | "3ArrowsGray"
        | "4ArrowsGray"
        | "5ArrowsGray"
        | "3Symbols"
        | "3Symbols2"
        | "3Flags";
    };

export type Image = {
  displayName: string;
  from: {
    col: number;
    row: number;
  };
  // In online Excel, when a bmp is inserted, it is converted a jpeg.
  // In online Excel, when a tiff is inserted, it is converted a png.
  extension: "png" | "jpeg" | "gif";
  data: Uint8Array;
  width: number;
  height: number;
};

export type ImageInfo = Omit<Image, "data"> & { fileBasename: string };

export type WorksheetOpts = {
  defaultColWidth?: number;
  defaultRowHeight?: number;
};

type RequiredWorksheetOpts = Required<WorksheetOpts>;

export type WorksheetType = {
  name: string;
  opts: RequiredWorksheetOpts;
  sheetData: SheetData;
  colOptsMap: Map<number, ColOpts>;
  rowOptsMap: Map<number, RowOpts>;
  mergeCells: MergeCell[];
  freezePane: FreezePane | null;
  mergeCellsModule: MergeCellsModule | null;
  conditionalFormattingModule: ConditionalFormattingModule | null;
  imageInfos: ImageInfo[];
  imageStore: ImageStore | null;
  imageModule: ImageModule | null;
  getCell(rowIndex: number, colIndex: number): NullableCell;
  setCell(rowIndex: number, colIndex: number, cell: SettableCell | null): void;
  setColOpts(col: ColOpts): void;
  setRowOpts(row: RowOpts): void;
  setFreezePane(freezePane: FreezePane): void;
};

/**
 * Standard Worksheet class
 */
export class Worksheet implements WorksheetType {
  private _name: string;
  private _opts: RequiredWorksheetOpts;
  private _sheetData: SheetData = [];
  private _colOptsMap = new Map<number, ColOpts>();
  private _rowOptsMap = new Map<number, RowOpts>();
  private _mergeCellsModule: MergeCellsModule = mergeCellsModule();
  private _freezePane: FreezePane | null = null;
  private _conditionalFormattingModule: ConditionalFormattingModule =
    conditionalFormattingModule();

  private _imageStore: ImageStore;
  private _imageModule: ImageModule = imageModule();

  constructor(
    name: string,
    imageStore?: ImageStore,
    opts: WorksheetOpts | undefined = {}
  ) {
    this._name = name;

    this._opts = {
      defaultColWidth: opts.defaultColWidth ?? DEFAULT_COL_WIDTH,
      defaultRowHeight: opts.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
    };

    this._imageStore = imageStore || new ImageStore();
  }

  get name() {
    return this._name;
  }

  get opts() {
    return this._opts;
  }

  set sheetData(sheetData: SheetData) {
    this._sheetData = sheetData;
  }

  get sheetData() {
    return this._sheetData;
  }

  get colOptsMap() {
    return this._colOptsMap;
  }

  get rowOptsMap() {
    return this._rowOptsMap;
  }

  get mergeCells() {
    return this._mergeCellsModule.getMergeCells();
  }

  get mergeCellsModule() {
    return this._mergeCellsModule;
  }

  get freezePane() {
    return this._freezePane;
  }

  get conditionalFormattings() {
    return this._conditionalFormattingModule.getConditionalFormattings();
  }

  get conditionalFormattingModule() {
    return this._conditionalFormattingModule;
  }

  get imageInfos() {
    return this._imageModule.getImageInfos();
  }

  get imageModule() {
    return this._imageModule;
  }

  get imageStore() {
    return this._imageStore;
  }

  getCell(rowIndex: number, colIndex: number): NullableCell {
    const rows = this._sheetData[rowIndex];
    if (!rows) {
      return null;
    }

    return rows[colIndex] || null;
  }

  // NOTE: `type: "merged"` is managed internally (blocked by setCell).
  // TODO: Prevent setting non-anchor cells inside merged ranges.
  setCell(rowIndex: number, colIndex: number, cell: SettableCell | null) {
    // Runtime guard for JS users / unsafe casts.
    if ((cell as any)?.type === "merged") {
      throw new Error(
        '`type: "merged"` is managed internally by mergeCellsModule. Use `setMergeCell()` instead.'
      );
    }

    if (!this._sheetData[rowIndex]) {
      const diff = rowIndex - this._sheetData.length + 1;
      for (let i = 0; i < diff; i++) {
        this._sheetData.push([]);
      }
    }

    const rows = this._sheetData[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push(null);
      }
    }

    rows[colIndex] = cell;
  }

  setColOpts(colOpts: ColOpts) {
    this._colOptsMap.set(colOpts.index, colOpts);
  }

  setRowOpts(rowOpts: RowOpts) {
    this._rowOptsMap.set(rowOpts.index, rowOpts);
  }

  setMergeCell(mergeCell: MergeCell) {
    this.mergeCellsModule.add(this, mergeCell);
  }

  setFreezePane(freezePane: FreezePane) {
    this._freezePane = freezePane;
  }

  setConditionalFormatting(conditionalFormatting: ConditionalFormatting) {
    this._conditionalFormattingModule.add(conditionalFormatting);
  }

  async insertImage(image: Image) {
    const fileBasename = await this._imageStore.addImage(
      image.data,
      image.extension
    );
    this._imageModule.add({ ...image, fileBasename });
  }
}

/**
 * Simplified Worksheet class
 */
export class WorksheetS implements WorksheetType {
  private _name: string;
  private _opts: RequiredWorksheetOpts;
  private _sheetData: SheetData = [];
  private _colOptsMap = new Map<number, ColOpts>();
  private _rowOptsMap = new Map<number, RowOpts>();
  private _mergeCellsModule = null;
  private _freezePane: FreezePane | null = null;
  private _conditionalFormattingModule = null;

  private _imageStore = null;
  private _imageModule = null;

  constructor(name: string, opts: WorksheetOpts | undefined = {}) {
    this._name = name;

    this._opts = {
      defaultColWidth: opts.defaultColWidth ?? DEFAULT_COL_WIDTH,
      defaultRowHeight: opts.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
    };
  }

  get name() {
    return this._name;
  }

  get opts() {
    return this._opts;
  }

  set sheetData(sheetData: SheetData) {
    this._sheetData = sheetData;
  }

  get sheetData() {
    return this._sheetData;
  }

  get colOptsMap() {
    return this._colOptsMap;
  }

  get rowOptsMap() {
    return this._rowOptsMap;
  }

  get mergeCells() {
    return [];
  }

  get mergeCellsModule() {
    return this._mergeCellsModule;
  }

  get freezePane() {
    return this._freezePane;
  }

  get conditionalFormattings() {
    return [];
  }

  get conditionalFormattingModule() {
    return this._conditionalFormattingModule;
  }

  get imageInfos() {
    return [];
  }

  get imageModule() {
    return this._imageModule;
  }

  get imageStore() {
    return this._imageStore;
  }

  getCell(rowIndex: number, colIndex: number): NullableCell {
    const rows = this._sheetData[rowIndex];
    if (!rows) {
      return null;
    }

    return rows[colIndex] || null;
  }

  // NOTE: `type: "merged"` is managed internally (blocked by setCell).
  // TODO: Prevent setting non-anchor cells inside merged ranges.
  setCell(rowIndex: number, colIndex: number, cell: SettableCell | null) {
    // Runtime guard for JS users / unsafe casts.
    if ((cell as any)?.type === "merged") {
      throw new Error(
        '`type: "merged"` is managed internally by mergeCellsModule. Use `setMergeCell()` instead.'
      );
    }

    if (!this._sheetData[rowIndex]) {
      const diff = rowIndex - this._sheetData.length + 1;
      for (let i = 0; i < diff; i++) {
        this._sheetData.push([]);
      }
    }

    const rows = this._sheetData[rowIndex]!;

    if (!rows[colIndex]) {
      const diff = colIndex - rows.length + 1;
      for (let i = 0; i < diff; i++) {
        rows.push(null);
      }
    }

    rows[colIndex] = cell;
  }

  setColOpts(colOpts: ColOpts) {
    this._colOptsMap.set(colOpts.index, colOpts);
  }

  setRowOpts(rowOpts: RowOpts) {
    this._rowOptsMap.set(rowOpts.index, rowOpts);
  }

  setFreezePane(freezePane: FreezePane) {
    this._freezePane = freezePane;
  }
}
