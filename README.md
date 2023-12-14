# xlsx-kaku

## Introduction

xlsx-kaku is a library for Node.js that outputs Excel xlsx files.  
It is exclusively for outputting xlsx files and cannot read them.

It currently only supports minimal functionality.
Please also see our [Roadmap](https://github.com/motinados/xlsx-kaku/issues/1).

## Example

```ts
async function main() {
  const wb = new Workbook();

  const ws = wb.addWorksheet("Sheet1");
  ws.setCell(0, 0, { type: "string", value: "Hello" });

  await wb.save("Hello.xlsx");
}
```
