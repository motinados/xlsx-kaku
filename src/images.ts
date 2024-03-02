let calcHashFn: (data: Uint8Array) => Promise<string>;

export async function getCalcHashFunc() {
  if (calcHashFn) {
    return calcHashFn;
  }

  let crypto: any;
  if (typeof window !== "undefined") {
    crypto = window.crypto;
  } else {
    crypto = await import("node:crypto");
  }

  calcHashFn = async (data: Uint8Array) => {
    const hashBuffer = await crypto.subtle.digest("SHA-256", data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const hashHex = hashArray
      .map((b) => b.toString(16).padStart(2, "0"))
      .join("");
    return hashHex;
  };

  return calcHashFn;
}
