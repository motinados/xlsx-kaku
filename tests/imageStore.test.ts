import { ImageStore } from "../src/imageStore";

describe("imageStore", () => {
  test("getHashFn", async () => {
    const imageStore = new ImageStore();
    const hashFn = await imageStore.getHashFn();
    expect(hashFn).toBeDefined();
  });

  test("addImage", async () => {
    const imageStore = new ImageStore();
    const extension = "png";

    const data1 = new Uint8Array([1, 2, 3]);
    const fileBasename = await imageStore.addImage(data1, extension);
    expect(fileBasename).toEqual("image1");
    expect(imageStore.getAllImages().size).toEqual(1);

    const data2 = new Uint8Array([4, 5, 6]);
    const fileBasename2 = await imageStore.addImage(data2, extension);
    expect(fileBasename2).toEqual("image2");
    expect(imageStore.getAllImages().size).toEqual(2);

    const data3 = new Uint8Array([1, 2, 3]);
    const fileBasename3 = await imageStore.addImage(data3, extension);
    expect(fileBasename3).toEqual("image1");
    expect(imageStore.getAllImages().size).toEqual(2);
  });

  test("addImage 2", async () => {
    const imageStore = new ImageStore();
    const data1 = new Uint8Array([1, 2, 3]);
    const data2 = new Uint8Array([4, 5, 6]);
    const data3 = new Uint8Array([1, 2, 3]);
    const extension = "png";

    const [b1, b2, b3] = await Promise.all([
      imageStore.addImage(data1, extension),
      imageStore.addImage(data2, extension),
      imageStore.addImage(data3, extension),
    ]);

    expect(imageStore.getAllImages().size).toEqual(2);
    expect(b1).toEqual(b3);
    expect(new Set([b1, b2])).toEqual(new Set(["image1", "image2"]));
  });

  test("getAllImages", async () => {
    const imageStore = new ImageStore();
    const data1 = new Uint8Array([1, 2, 3]);
    const extension = "png";
    await imageStore.addImage(data1, extension);
    expect(imageStore.getAllImages().size).toEqual(1);

    const data2 = new Uint8Array([4, 5, 6]);
    await imageStore.addImage(data2, extension);
    expect(imageStore.getAllImages().size).toEqual(2);

    const data3 = new Uint8Array([7, 8, 9]);
    await imageStore.addImage(data3, extension);
    expect(imageStore.getAllImages().size).toEqual(3);
  });
});
