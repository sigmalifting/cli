const createUuidFromBytes = (bytes) => {
  const byteToHex = (value) => value.toString(16).padStart(2, "0");
  const segments = [
    Array.from(bytes.slice(0, 4), byteToHex).join(""),
    Array.from(bytes.slice(4, 6), byteToHex).join(""),
    Array.from(bytes.slice(6, 8), byteToHex).join(""),
    Array.from(bytes.slice(8, 10), byteToHex).join(""),
    Array.from(bytes.slice(10, 16), byteToHex).join(""),
  ];

  return segments.join("-");
};

const generateUuid = () => {
  if (typeof globalThis.crypto?.randomUUID === "function") {
    return globalThis.crypto.randomUUID();
  }

  const bytes = new Uint8Array(16);

  if (typeof globalThis.crypto?.getRandomValues === "function") {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    for (let index = 0; index < bytes.length; index += 1) {
      bytes[index] = Math.floor(Math.random() * 256);
    }
  }

  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  return createUuidFromBytes(bytes);
};

export const generateId = (prefix = "") => {
  const uuid = generateUuid();
  return prefix ? `${prefix}_${uuid}` : uuid;
};
