import { stringifySorted } from "./fonts";

// https://learn.microsoft.com/ja-jp/dotnet/api/documentformat.openxml.spreadsheet.borderstylevalues?view=openxml-2.8.1
type BorderStyle =
  | "dashDot"
  | "dashDotDot"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "medium"
  | "mediumDashDot"
  | "mediumDashDotDot"
  | "mediumDashed"
  | "none"
  | "slantDashDot"
  | "thick"
  | "thin";

export type Border = {
  left?: {
    style: BorderStyle;
    color: string;
  };
  right?: {
    style: BorderStyle;
    color: string;
  };
  top?: {
    style: BorderStyle;
    color: string;
  };
  bottom?: {
    style: BorderStyle;
    color: string;
  };
  diagonal?: {
    style: BorderStyle;
    color: string;
  };
};

export class Borders {
  // default border
  //   <borders count="1">
  //       <border>
  //           <left/>
  //           <right/>
  //           <top/>
  //           <bottom/>
  //           <diagonal/>
  //       </border>
  //   </borders>
  private borders = new Map<string, number>();

  getBorderId(border: Border): number {
    if (Object.keys(border).length === 0) {
      return 0;
    }

    const key = stringifySorted(border);
    const id = this.borders.get(key);
    if (id !== undefined) {
      return id;
    }

    // boders map does not have a default border, so retuns size + 1
    const borderId = this.borders.size + 1;
    this.borders.set(key, borderId);
    return borderId;
  }
}
