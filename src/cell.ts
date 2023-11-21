export class Cell {
  private _type: "string" | "number" = "string";
  private _value: string | number = "";
  constructor() {}

  get type() {
    return this._type;
  }

  set value(value: string | number) {
    if (typeof value === "number") {
      this._type = "number";
    } else {
      this._type = "string";
    }
    this._value = value;
  }

  get value() {
    return this._value;
  }
}
