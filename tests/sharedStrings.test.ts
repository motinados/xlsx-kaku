import { SharedStrings } from "../src/sharedStrings";
describe("SharedStrings", () => {
  test("should be able to create a sharedStrings", () => {
    const sharedStrings = new SharedStrings();
    expect(sharedStrings).toBeInstanceOf(SharedStrings);
  });

  test("should be able to get index", () => {
    const sharedStrings = new SharedStrings();
    expect(sharedStrings.getIndex("hello")).toBe(0);
    expect(sharedStrings.getIndex("world")).toBe(1);
    expect(sharedStrings.getIndex("hello")).toBe(0);
  });

  test("should be able to get values in order", () => {
    const sharedStrings = new SharedStrings();
    sharedStrings.getIndex("hello");
    sharedStrings.getIndex("world");
    sharedStrings.getIndex("hello");
    expect(sharedStrings.getValuesInOrder()).toEqual(["hello", "world"]);
  });
});
