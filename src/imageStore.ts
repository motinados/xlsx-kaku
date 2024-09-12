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

  private async createHashFn() {
    let crypto: any;
    if (typeof window !== "undefined") {
      crypto = window.crypto;
    } else {
      crypto = await import("node:crypto");
    }

    return async (data: Uint8Array) => {
      const hashBuffer = await crypto.subtle.digest("SHA-256", data);
      const hashArray = Array.from(new Uint8Array(hashBuffer));
      const hashHex = hashArray
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");
      return hashHex;
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
