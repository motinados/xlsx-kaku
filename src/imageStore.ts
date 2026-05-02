type HashFn = (data: Uint8Array) => Promise<string>;

type ImageData = {
  fileBasename: string;
  extension: string;
  data: Uint8Array;
};

export class ImageStore {
  private images: Map<string, ImageData> = new Map();
  private _hashFn: HashFn | null = null;
  private _initializationPromise: Promise<void> | null = null;

  constructor() {}

  private async initHashFn() {
    this._hashFn = await this.createHashFn();
  }

  private toArrayBuffer(data: Uint8Array): ArrayBuffer {
    const buffer = new ArrayBuffer(data.byteLength);
    new Uint8Array(buffer).set(data);
    return buffer;
  }

  private async createHashFn(): Promise<HashFn> {
    const cryptoObj = globalThis.crypto;

    if (!cryptoObj?.subtle) {
      throw new Error(
        "ImageStore requires Web Crypto API: globalThis.crypto.subtle is not available."
      );
    }

    return async (data: Uint8Array) => {
      const hashBuffer = await cryptoObj.subtle.digest(
        "SHA-256",
        this.toArrayBuffer(data)
      );
      const hashArray = Array.from(new Uint8Array(hashBuffer));

      return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
    };
  }

  async getHashFn(): Promise<HashFn> {
    if (this._hashFn) {
      return this._hashFn;
    }

    if (!this._initializationPromise) {
      this._initializationPromise = this.initHashFn();
    }

    await this._initializationPromise;
    return this._hashFn!;
  }

  async addImage(data: Uint8Array, extension: string) {
    const hashFn = await this.getHashFn();
    const hash = await hashFn(data);

    if (this.images.has(hash)) {
      return this.images.get(hash)!.fileBasename;
    }

    const fileBasename = `image${this.images.size + 1}`;
    this.images.set(hash, { fileBasename, data, extension });
    return fileBasename;
  }

  getAllImages() {
    return this.images;
  }
}
