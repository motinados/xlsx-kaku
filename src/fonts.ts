type Font = {
  name: string;
  size: number;
  color: string;
  family?: number;
  scheme?: string;
};

const fonts = new Map<string, number>([
  [
    stringifySorted({
      name: "游ゴシック",
      size: 11,
      color: "FF0000",
    }),
    0,
  ],
]);

export function getFontId(font: Font): number {
  const key = stringifySorted(font);
  const id = fonts.get(key);
  if (id !== undefined) {
    return id;
  }

  const fontId = fonts.size;
  fonts.set(key, fontId);
  return fontId;
}

function sortObjectKeys(obj: Record<string, any>): Record<string, any> {
  return Object.keys(obj)
    .sort()
    .reduce((sortedObj, key) => {
      sortedObj[key] = obj[key];
      return sortedObj;
    }, {} as Record<string, any>);
}

export function stringifySorted(obj: Record<string, any>): string {
  const sortedObj = sortObjectKeys(obj);
  return JSON.stringify(sortedObj);
}
