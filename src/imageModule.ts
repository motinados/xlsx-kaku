import { DrawingRels } from "./drawingRels";
import { Image } from "./worksheet";
import { XlsxImage } from "./xml/worksheetXml";

export type ImageModule = {
  name: string;
  getImages(): Image[];
  createXlsxImage(image: Image, drawingRels: DrawingRels): XlsxImage;
};

export function imageModule(): ImageModule {
  const images: Image[] = [];
  return {
    name: "image",
    getImages() {
      return images;
    },
    createXlsxImage(image: Image, drawingRels: DrawingRels): XlsxImage {
      const num = drawingRels.length + 1;
      const rId = drawingRels.addDrawingRel({
        target: `../media/image${num}.png`,
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      });

      return {
        rId,
        // Files created in online Excel seem to start with a sequential number begginning from 2.
        id: String(num + 1),
        name: image.displayName,
        editAs: "oneCell",
        from: {
          col: image.from.col,
          colOff: 0,
          row: image.from.row,
          rowOff: 0,
        },
        ext: {
          cx: (914400 / 96) * image.width,
          cy: (914400 / 96) * image.height,
        },
      };
    },
  };
}
