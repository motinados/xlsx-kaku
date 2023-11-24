export class SharedStrings {
  private map: Map<string, number>;

  constructor() {
    this.map = new Map();
  }

  getIndex(value: string): number {
    if (this.map.has(value)) {
      return this.map.get(value)!;
    }

    const newIndex = this.map.size;
    this.map.set(value, newIndex);

    return newIndex;
  }
}
