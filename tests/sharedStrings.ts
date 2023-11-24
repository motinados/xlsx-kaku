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
});
