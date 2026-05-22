import { Dxf } from "../dxf";
import { ConditionalFormatting } from "../worksheet";
import { XlsxConditionalFormatting } from "../xml/worksheetXml";
import { createXlsxConditionalFormattings } from "./conditionalFormattingConverter";
import { makeConditionalFormattingXml } from "./conditionalFormattingXml";

export type ConditionalFormattingModule = {
  name: string;
  getConditionalFormattings(): ConditionalFormatting[];
  add(conditionalFormatting: ConditionalFormatting): void;
  createXlsxConditionalFormatting(
    conditionalFormattings: ConditionalFormatting[],
    dxf: Dxf
  ): XlsxConditionalFormatting[];
  makeXmlElm(formattings: XlsxConditionalFormatting[]): string;
};

export function conditionalFormattingModule(): ConditionalFormattingModule {
  const conditionalFormattings: ConditionalFormatting[] = [];
  return {
    name: "conditional-formatting",
    getConditionalFormattings() {
      return conditionalFormattings;
    },
    add(conditionalFormatting: ConditionalFormatting) {
      conditionalFormattings.push(conditionalFormatting);
    },
    createXlsxConditionalFormatting(
      conditionalFormattings: ConditionalFormatting[],
      dxf: Dxf
    ) {
      return createXlsxConditionalFormattings(conditionalFormattings, dxf);
    },
    makeXmlElm(formattings: XlsxConditionalFormatting[]) {
      return makeConditionalFormattingXml(formattings);
    },
  };
}
