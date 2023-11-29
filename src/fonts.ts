type Font = {
  name: string;
  size: number;
  color: string;
  family?: number;
  scheme?: string;
};

export class Fonts {
  private fonts = new Map<string, number>([
    [
      stringifySorted({
        name: "游ゴシック",
        size: 11,
        color: "FF0000",
      }),
      0,
    ],
  ]);

  getFontId(font: Font): number {
    const key = stringifySorted(font);
    const id = this.fonts.get(key);
    if (id !== undefined) {
      return id;
    }

    const fontId = this.fonts.size;
    this.fonts.set(key, fontId);
    return fontId;
  }
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
