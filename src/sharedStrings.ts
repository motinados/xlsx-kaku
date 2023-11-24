export class SharedStrings {
  private map: Map<string, number>;
  private _count: number = 0;

  constructor() {
    this.map = new Map();
  }

  get count() {
    return this._count;
  }

  get uniqueCount() {
    return this.map.size;
  }

  getIndex(value: string): number {
    this._count++;

    if (this.map.has(value)) {
      return this.map.get(value)!;
    }

    const newIndex = this.map.size;
    this.map.set(value, newIndex);

    return newIndex;
  }

  getValuesInOrder(): string[] {
    const sortedEntries = Array.from(this.map.entries()).sort(
      (a, b) => a[1] - b[1]
    );
    return sortedEntries.map((entry) => entry[0]);
  }
}
