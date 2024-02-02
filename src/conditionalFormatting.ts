import { DxfStyle } from "./dxf";

export type ConditionalFormatting = {
  type: "top10";
  sqref: string;
  dxfId: number;
  priority: number;
  percent: boolean;
  rank: number;
};

export type ConditionalFormattingProps = {
  type: "top10";
  sqref: string;
  priority: number;
  percent: boolean;
  rank: number;
  style: DxfStyle;
};
