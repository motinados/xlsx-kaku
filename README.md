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

### changing the width of columns

```ts
import { Workbook } from "xlsx-kaku";

async function main() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  // The width of only column A willl be changed.
  ws.setColWidth({ min: 1, max: 1, width: 12 });

  // The width of columns B to F will be changed.
  ws.setColWidth({ min: 2, max: 6, width: 24 });

  await wb.save("test.xlsx");
}
```
