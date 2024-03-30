import {
  ConditionalFormattingModule,
  conditionalFormattingModule,
} from "./conditionalFormattingModule";
import { DxfStyle } from "./dxf";
import { ImageModule, imageModule } from "./imageModule";
import { ImageStore } from "./imageStore";
import { MergeCellsModule, mergeCellsModule } from "./mergeCellsModule";
import { CellStyle, NullableCell, SheetData } from "./sheetData";

/**
 * The value is the same as the one in files created with Online Excel.
 * Changing this value will result in differences in integration tests.
 */
export const DEFAULT_COL_WIDTH = 9;
export const DEFAULT_ROW_HEIGHT = 13.5;

export type ColProps = {
  index: number;
  width?: number;
  style?: CellStyle;
};

export type RowProps = {
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
  // TODO: Support other image formats.
  // In online Excel, when a bmp is inserted, it is converted a jpeg.
  // In online Excel, when a tiff is inserted, it is converted a png.
  extension: "png" | "jpeg" | "gif";
  data: Uint8Array;
  width: number;
  height: number;
};

export type ImageInfo = Omit<Image, "data">;

export type WorksheetProps = {
  defaultColWidth?: number;
  defaultRowHeight?: number;
};

type RequiredWorksheetProps = Required<WorksheetProps>;

export type WorksheetType = {
  name: string;
  props: RequiredWorksheetProps;
  sheetData: SheetData;
  cols: Map<number, ColProps>;
  rows: Map<number, RowProps>;
  mergeCells: MergeCell[];
  freezePane: FreezePane | null;
  mergeCellsModule: MergeCellsModule | null;
  conditionalFormattingModule: ConditionalFormattingModule | null;
  imageInfos: ImageInfo[];
  imageStore: ImageStore | null;
  imageModule: ImageModule | null;
  getCell(rowIndex: number, colIndex: number): NullableCell;
  setCell(rowIndex: number, colIndex: number, cell: NullableCell): void;
  setColProps(col: ColProps): void;
  setRowProps(row: RowProps): void;
  setFreezePane(freezePane: FreezePane): void;
};

/**
 * Standard Worksheet class
 */
export class Worksheet implements WorksheetType {
  private _name: string;
  private _props: RequiredWorksheetProps;
  private _sheetData: SheetData = [];
  private _cols = new Map<number, ColProps>();
  private _rows = new Map<number, RowProps>();
  private _mergeCellsModule: MergeCellsModule = mergeCellsModule();
  private _freezePane: FreezePane | null = null;
  private _conditionalFormattingModule: ConditionalFormattingModule =
    conditionalFormattingModule();

  private _imageStore: ImageStore;
  private _imageModule: ImageModule = imageModule();

  constructor(
    name: string,
    imageStore?: ImageStore,
    props: WorksheetProps | undefined = {}
  ) {
    this._name = name;

    this._props = {
      defaultColWidth: props.defaultColWidth ?? DEFAULT_COL_WIDTH,
      defaultRowHeight: props.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
    };

    this._imageStore = imageStore || new ImageStore();
  }

  get name() {
    return this._name;
  }

  get props() {
    return this._props;
  }

  set sheetData(sheetData: SheetData) {
    this._sheetData = sheetData;
  }

  get sheetData() {
    return this._sheetData;
  }

  get cols() {
    return this._cols;
  }

  get rows() {
    return this._rows;
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

  // TODO: Cells that have been merged cannot be set.
  setCell(rowIndex: number, colIndex: number, cell: NullableCell) {
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

  setColProps(colProps: ColProps) {
    this._cols.set(colProps.index, colProps);
  }

  setRowProps(row: RowProps) {
    this._rows.set(row.index, row);
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
    await this._imageStore.addImage(image.data, image.extension);
    this._imageModule.add(image);
  }
}

/**
 * Simplified Worksheet class
 */
export class WorksheetS implements WorksheetType {
  private _name: string;
  private _props: RequiredWorksheetProps;
  private _sheetData: SheetData = [];
  private _cols = new Map<number, ColProps>();
  private _rows = new Map<number, RowProps>();
  private _mergeCellsModule = null;
  private _freezePane: FreezePane | null = null;
  private _conditionalFormattingModule = null;

  private _imageStore = null;
  private _imageModule = null;

  constructor(name: string, props: WorksheetProps | undefined = {}) {
    this._name = name;

    this._props = {
      defaultColWidth: props.defaultColWidth ?? DEFAULT_COL_WIDTH,
      defaultRowHeight: props.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
    };
  }

  get name() {
    return this._name;
  }

  get props() {
    return this._props;
  }

  set sheetData(sheetData: SheetData) {
    this._sheetData = sheetData;
  }

  get sheetData() {
    return this._sheetData;
  }

  get cols() {
    return this._cols;
  }

  get rows() {
    return this._rows;
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

  // TODO: Cells that have been merged cannot be set.
  setCell(rowIndex: number, colIndex: number, cell: NullableCell) {
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

  setColProps(colProps: ColProps) {
    this._cols.set(colProps.index, colProps);
  }

  setRowProps(row: RowProps) {
    this._rows.set(row.index, row);
  }

  setFreezePane(freezePane: FreezePane) {
    this._freezePane = freezePane;
  }
}
