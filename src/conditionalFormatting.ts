import { DxfStyle } from "./dxf";

export type ConditionalFormattingProps = {
  type: "top10";
  sqref: string;
  priority: number;
  percent: boolean;
  rank: number;
  style: DxfStyle;
};
