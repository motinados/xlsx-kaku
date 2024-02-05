# xlsx-kaku

## Introduction

xlsx-kaku is a library for Node.js that outputs Excel xlsx files.  
It is exclusively for outputting xlsx files and cannot read them.

It currently supports basic functionality.
Please also see our [Roadmap](https://github.com/motinados/xlsx-kaku/issues/1).

> This library is currently in the early stages of development.
> We are constantly working to improve and optimize our codebase, which may lead to some breaking changes.
> We recommend regularly checking the latest release logs and documentation.

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

  const xlsx = wb.generateXlsxSync();
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

    const xlsx = wb.generateXlsxSync();

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

### Worksheet

```ts
import { Workbook } from "xlsx-kaku";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

// Set default coloum width and default row height
const ws = wb.addWorksheet("Sheet2", {
  defaultColWidth: 30,
  defaultRowHeight: 16,
});
```

### Cell

```ts
import { Workbook } from "xlsx-kaku";

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
  linkType: "external",
  text: "google",
  value: "https://www.google.com/",
});
ws.setCell(3, 1, {
  type: "hyperlink",
  linkType: "internal",
  text: "to A1",
  value: "Sheet1!A1",
});
ws.setCell(3, 2, {
  type: "hyperlink",
  linkType: "email",
  text: "sample",
  value: "sample@example.com",
});

// boolean
ws.setCell(4, 0, { type: "boolean", value: true });
ws.setCell(4, 1, { type: "boolean", value: false });

// formula
ws.setCell(5, 0, { type: "number", value: 1 });
ws.setCell(5, 1, { type: "number", value: 2 });
ws.setCell(5, 2, { type: "formula", value: "SUM(A6:B6)" });

const xlsx = wb.generateXlsxSync();
```

### Column

```ts
import { Workbook } from "xlsx-kaku";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

ws.setColProps({
  index: 0,
  width: 12,
  style: { fill: { patternType: "solid", fgColor: "FFFFFF00" } },
});

const xlsx = wb.generateXlsxSync();
```

### Row

```ts
import { Workbook } from "xlsx-kaku";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

// Currently, RowProps do not work without a value in the cell.
ws.setCell(0, 0, { type: "number", value: 1 });
ws.setRowProps({
  index: 0,
  height: 20,
  style: { fill: { patternType: "solid", fgColor: "FFFFFF00" } },
});

const xlsx = wb.generateXlsxSync();
```

### Alignment

```ts
import { Workbook } from "xlsx-kaku";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

ws.setCell(0, 0, {
  type: "number",
  value: 12,
  style: {
    alignment: {
      horizontal: "center",
      vertical: "top",
      textRotation: 90,
      wordwrap: true,
    },
  },
});

const xlsx = wb.generateXlsxSync();
```

### Merge cells

```ts
import { Workbook } from "xlsx-kaku";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

ws.setCell(0, 0, { type: "number", value: 1 });
ws.setCell(1, 0, { type: "number", value: 2 });
ws.setMergeCell({ ref: "A1:C1" });
ws.setMergeCell({ ref: "A2:A4" });

const xlsx = wb.generateXlsxSync();
```

### Freeze pane

```ts
import { Workbook } from "xlsx-kaku";

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

const xlsx = wb.generateXlsxSync();
```

### Conditional Formatting

```ts
const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

// The following types are supported for numbers:
//
// "top", "bottom", "aboveAverage", "belowAverage", "atOrAboveAverage", "atOrBelowAverage",
// "duplicateValues", "greaterThan", "lessThan", "equal", "between"
//
// Here, we present example of "top".
ws.setCell(0, 0, { type: "number", value: 1 });
ws.setCell(1, 0, { type: "number", value: 2 });
ws.setCell(2, 0, { type: "number", value: 3 });
ws.setCell(3, 0, { type: "number", value: 4 });
ws.setCell(4, 0, { type: "number", value: 5 });
ws.setCell(5, 0, { type: "number", value: 6 });
ws.setCell(6, 0, { type: "number", value: 7 });
ws.setCell(7, 0, { type: "number", value: 8 });
ws.setCell(8, 0, { type: "number", value: 9 });
ws.setCell(9, 0, { type: "number", value: 10 });

ws.setConditionalFormatting({
  sqref: "A:A",
  type: "top",
  priority: 1,
  percent: true,
  rank: 10,
  style: {
    font: { color: "FF9C0006" },
    fill: { bgColor: "FFFFC7CE" },
  },
});

// The following types are supported for strings:
//
// "containsText", "notContainsText", "beginsWith", "endsWith";
//
// Here, we present example of "containsText"
ws.setCell(0, 1, { type: "string", value: "Lion" });
ws.setCell(1, 1, { type: "string", value: "Elephant" });
ws.setCell(2, 1, { type: "string", value: "Panda" });
ws.setCell(3, 1, { type: "string", value: "Giraffe" });
ws.setCell(4, 1, { type: "string", value: "Turtle" });
ws.setCell(5, 1, { type: "string", value: "Apple" });
ws.setCell(6, 1, { type: "string", value: "Banana" });
ws.setCell(7, 1, { type: "string", value: "Strawberry" });
ws.setCell(8, 1, { type: "string", value: "Melon" });
ws.setCell(9, 1, { type: "string", value: "Orange" });

ws.setConditionalFormatting({
  type: "containsText",
  sqref: "B1:B1048576",
  text: "a",
  style: {
    font: { color: "FF9C0006" },
    fill: { bgColor: "FFFFC7CE" },
  },
  priority: 1,
});

const xlsx = wb.generateXlsxSync();
writeFileSync("test.xlsx", xlsx);
```
