import { stringifySorted } from "./fonts";

type Fill =
  | {
      patternType: "none";
    }
  | {
      patternType: "gray125";
    }
  | {
      patternType: "solid";
      fgColor: string;
    };

// <fills count="3">
//     <fill>
//         <patternFill patternType="none"/>
//     </fill>
//     <fill>
//         <patternFill patternType="gray125"/>
//     </fill>
//     <fill>
//         <patternFill patternType="solid">
//             <fgColor rgb="FFFFFF00"/>
//             <bgColor indexed="64"/>
//         </patternFill>
//     </fill>
// </fills>
export class Fills {
  private fills = new Map<string, number>([
    [
      stringifySorted({
        patternType: "none",
      }),
      0,
    ],
    [
      stringifySorted({
        patternType: "gray125",
      }),
      1,
    ],
  ]);

  getFillId(fill: Fill): number {
    const key = stringifySorted(fill);
    const id = this.fills.get(key);
    if (id !== undefined) {
      return id;
    }

    const fillId = this.fills.size;
    this.fills.set(key, fillId);
    return fillId;
  }

  makeXml(): string {
    let xml = `<fills count="${this.fills.size}">`;
    this.fills.forEach((_, key) => {
      const fill = JSON.parse(key) as Fill;

      xml += "<fill>";
      if (fill.patternType === "none" || fill.patternType === "gray125") {
        xml += `<patternFill patternType="${fill.patternType}"/>`;
      } else {
        xml += `<patternFill patternType="${fill.patternType}">`;
        xml += `<fgColor rgb="${fill.fgColor}"/>`;
        xml += '<bgColor indexed="64"/>';
        xml += "</patternFill>";
      }
      xml += "</fill>";
    });
    xml += "</fills>";
    return xml;
  }
}
