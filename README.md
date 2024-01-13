# xlsx-kaku

## Introduction

xlsx-kaku is a library for Node.js that outputs Excel xlsx files.  
It is exclusively for outputting xlsx files and cannot read them.

It currently only supports minimal functionality.
Please also see our [Roadmap](https://github.com/motinados/xlsx-kaku/issues/1).

## Installation

```
npm install xlsx-kaku
```

## Example

### Basic Usage on the Server-Side

```ts
import { writeFileSync } from "node:fs";
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, { type: "string", value: "Hello" });
  ws.setCell(0, 1, { type: "number", value: 123 });

  const xlsx = wb.generateXlsx();
  writeFileSync("sample.xlsx", xlsx);
}
```

### Basic Usage on the Client-Side (React)

```ts
import { Workbook } from "xlsx-kaku";

export default function DownloadButton() {
  const handleDownload = () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.setCell(0, 0, { type: "string", value: "Hello" });
    ws.setCell(0, 1, { type: "number", value: 123 });

    const xlsx = wb.generateXlsx();

    const blob = new Blob([xlsx], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "sample.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      <button onClick={handleDownload}>download xlsx</button>
    </div>
  );
}
```

### Cell

```ts
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  // string
  ws.setCell(0, 0, { type: "string", value: "Hello" });
  ws.setCell(0, 1, { type: "string", value: "World" });

  // number
  ws.setCell(1, 0, { type: "number", value: 1 });
  ws.setCell(1, 1, { type: "number", value: 2 });

  // date
  ws.setCell(2, 0, {
    type: "date",
    value: new Date().toISOString(),
    style: { numberFormat: { formatCode: "yyyy-mm-dd" } },
  });

  // hyperlink
  ws.setCell(3, 0, {
    type: "hyperlink",
    value: "https://www.google.com",
  });

  // boolean
  ws.setCell(4, 0, { type: "boolean", value: true });
  ws.setCell(4, 1, { type: "boolean", value: false });

  // formula
  ws.setCell(5, 0, { type: "number", value: 1 });
  ws.setCell(5, 1, { type: "number", value: 2 });
  ws.setCell(5, 2, { type: "formula", value: "SUM(A6:B6)" });

  const xlsx = wb.generateXlsx();
}
```

### Column

```ts
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  // The width of only column A will be changed.
  ws.setColWidth({ startIndex: 0, endIndex: 0, width: 12 });

  // The width of columns B to C will be changed.
  ws.setColWidth({ startIndex: 1, endIndex: 2, width: 24 });

  // The style of column D will be changed.
  ws.setColStyle({
    startIndex: 3,
    endIndex: 3,
    style: { fill: { patternType: "solid", fgColor: "FFFFFF00" } },
  });

  const xlsx = wb.generateXlsx();
}
```

### Row

```ts
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, { type: "number", value: 1 });
  ws.setCell(1, 0, { type: "number", value: 2 });
  ws.setCell(2, 0, { type: "number", value: 3 });
  ws.setCell(3, 0, { type: "number", value: 4 });

  // Change the height of the first row.
  ws.setRowHeight({ index: 0, height: 20.25 });

  // Unlike column width, it is necessary to set each row individually.
  ws.setRowHeight({ index: 1, height: 39.75 });
  ws.setRowHeight({ index: 2, height: 39.75 });
  ws.setRowHeight({ index: 3, height: 39.75 });

  // Change the color of the third row.
  ws.setRowStyle({
    index: 2,
    style: { fill: { patternType: "solid", fgColor: "FFFFFF00" } },
  });

  const xlsx = wb.generateXlsx();
}
```

### Alignment

```ts
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, {
    type: "number",
    value: 12,
    style: {
      alignment: { horizontal: "center", vertical: "top", textRotation: 90 },
    },
  });

  const xlsx = wb.generateXlsx();
}
```

### Merge cells

```ts
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, { type: "number", value: 1 });
  ws.setCell(1, 0, { type: "number", value: 2 });
  ws.setMergeCell({ ref: "A1:C1" });
  ws.setMergeCell({ ref: "A2:A4" });

  const xlsx = wb.generateXlsx();
}
```

### Freeze pane

```ts
import { Workbook } from "xlsx-kaku";

function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, { type: "number", value: 1 });
  ws.setCell(0, 1, { type: "number", value: 2 });
  ws.setCell(1, 0, { type: "number", value: 3 });
  ws.setCell(1, 1, { type: "number", value: 4 });

  // the first row will be fixed.
  ws.setFreezePane({ target: "row", split: 1 });

  // Column A will be fixed.
  // ws.setFreezePane({ target: "column", split: 1 });

  const xlsx = wb.generateXlsx();
}
```
