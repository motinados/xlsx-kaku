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

### basic usage

```ts
import { Workbook } from "xlsx-kaku";

async function main() {
  const wb = new Workbook();

  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, { type: "string", value: "Hello" });
  ws.setCell(0, 1, { type: "number", value: 123 });
  ws.setCell(1, 0, {
    type: "date",
    value: new Date().toISOString(),
    style: {
      numberFormat: { formatCode: "yyyy-mm-dd" },
    },
  });

  await wb.save("Hello.xlsx");
}
```

### merge cells

```ts
import { Workbook } from "xlsx-kaku";

async function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.setCell(0, 0, { type: "number", value: 1 });
  ws.setCell(1, 0, { type: "number", value: 2 });
  ws.setMergeCell({ ref: "A1:C1" });
  ws.setMergeCell({ ref: "A2:A4" });

  await wb.save("test.xlsx");
}
```

### changing the width of columns

```ts
import { Workbook } from "xlsx-kaku";

async function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  // The width of only column A will be changed.
  ws.setColWidth({ min: 1, max: 1, width: 12 });

  // The width of columns B to F will be changed.
  ws.setColWidth({ min: 2, max: 6, width: 24 });

  await wb.save("test.xlsx");
}
```

### changing the height of rows

```ts
import { Workbook } from "xlsx-kaku";

async function main() {
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

  await wb.save("test.xlsx");
}
```
