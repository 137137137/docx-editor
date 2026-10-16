// Synchronous platform-neutral SHA-256 over raw bytes.
//
// Lives in the store lane because both trust boundaries that fingerprint bytes sit here:
// media parts (image resources) and embedded fonts hash their payloads before anything
// downstream consumes them. The layout lane re-exports `sha256FontBytes` from
// `font-resource.ts`, which is the direction the lane DAG allows (layout → store).

const SHA256_INITIAL = new Uint32Array([
  0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
]);

const SHA256_CONSTANTS = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

const rotateRight = (value: number, count: number): number =>
  (value >>> count) | (value << (32 - count));

/** Pure-JavaScript SHA-256, the reference every runtime agrees with. Exported for tests. */
export const sha256FontBytesPure = (bytes: Uint8Array): string => {
  const paddedLength = Math.ceil((bytes.byteLength + 9) / 64) * 64;
  const padded = new Uint8Array(paddedLength);
  padded.set(bytes);
  padded[bytes.byteLength] = 0x80;
  const bitLength = BigInt(bytes.byteLength) * 8n;
  for (let index = 0; index < 8; index += 1) {
    padded[paddedLength - 1 - index] = Number((bitLength >> BigInt(index * 8)) & 0xffn);
  }

  const hash = SHA256_INITIAL.slice();
  const words = new Uint32Array(64);
  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let index = 0; index < 16; index += 1) {
      const start = offset + index * 4;
      words[index] =
        ((padded[start]! << 24) |
          (padded[start + 1]! << 16) |
          (padded[start + 2]! << 8) |
          padded[start + 3]!) >>>
        0;
    }
    for (let index = 16; index < 64; index += 1) {
      const left = words[index - 15]!;
      const right = words[index - 2]!;
      const sigma0 = rotateRight(left, 7) ^ rotateRight(left, 18) ^ (left >>> 3);
      const sigma1 = rotateRight(right, 17) ^ rotateRight(right, 19) ^ (right >>> 10);
      words[index] = (words[index - 16]! + sigma0 + words[index - 7]! + sigma1) >>> 0;
    }

    let a = hash[0]!;
    let b = hash[1]!;
    let c = hash[2]!;
    let d = hash[3]!;
    let e = hash[4]!;
    let f = hash[5]!;
    let g = hash[6]!;
    let h = hash[7]!;
    for (let index = 0; index < 64; index += 1) {
      const sum1 = rotateRight(e, 6) ^ rotateRight(e, 11) ^ rotateRight(e, 25);
      const choice = (e & f) ^ (~e & g);
      const temp1 = (h + sum1 + choice + SHA256_CONSTANTS[index]! + words[index]!) >>> 0;
      const sum0 = rotateRight(a, 2) ^ rotateRight(a, 13) ^ rotateRight(a, 22);
      const majority = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (sum0 + majority) >>> 0;
      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }
    hash[0] = (hash[0]! + a) >>> 0;
    hash[1] = (hash[1]! + b) >>> 0;
    hash[2] = (hash[2]! + c) >>> 0;
    hash[3] = (hash[3]! + d) >>> 0;
    hash[4] = (hash[4]! + e) >>> 0;
    hash[5] = (hash[5]! + f) >>> 0;
    hash[6] = (hash[6]! + g) >>> 0;
    hash[7] = (hash[7]! + h) >>> 0;
  }
  return `sha256:${Array.from(hash, (word) => word.toString(16).padStart(8, '0')).join('')}`;
};

/**
 * The compression function over whole 64-byte blocks of `block`, from `start` to `end`, into
 * `state`. Shared by the fast paths below; the reference above stays as written.
 */
const blockWords = new Int32Array(64);
const hexOfByte = Array.from({ length: 256 }, (_, byte) => byte.toString(16).padStart(2, '0'));

function compressBlocks(state: Int32Array, block: Uint8Array, start: number, end: number): void {
  const words = blockWords;
  const k = SHA256_CONSTANTS;
  let h0 = state[0]!;
  let h1 = state[1]!;
  let h2 = state[2]!;
  let h3 = state[3]!;
  let h4 = state[4]!;
  let h5 = state[5]!;
  let h6 = state[6]!;
  let h7 = state[7]!;
  for (let offset = start; offset < end; offset += 64) {
    for (let index = 0; index < 16; index += 1) {
      const at = offset + index * 4;
      words[index] =
        (block[at]! << 24) | (block[at + 1]! << 16) | (block[at + 2]! << 8) | block[at + 3]!;
    }
    for (let index = 16; index < 64; index += 1) {
      const left = words[index - 15]!;
      const right = words[index - 2]!;
      const sigma0 = ((left >>> 7) | (left << 25)) ^ ((left >>> 18) | (left << 14)) ^ (left >>> 3);
      const sigma1 =
        ((right >>> 17) | (right << 15)) ^ ((right >>> 19) | (right << 13)) ^ (right >>> 10);
      words[index] = (words[index - 16]! + sigma0 + words[index - 7]! + sigma1) | 0;
    }
    let a = h0;
    let b = h1;
    let c = h2;
    let d = h3;
    let e = h4;
    let f = h5;
    let g = h6;
    let h = h7;
    for (let index = 0; index < 64; index += 1) {
      const sum1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const choice = (e & f) ^ (~e & g);
      const temp1 = (h + sum1 + choice + k[index]! + words[index]!) | 0;
      const sum0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const majority = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (sum0 + majority) | 0;
      h = g;
      g = f;
      f = e;
      e = (d + temp1) | 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) | 0;
    }
    h0 = (h0 + a) | 0;
    h1 = (h1 + b) | 0;
    h2 = (h2 + c) | 0;
    h3 = (h3 + d) | 0;
    h4 = (h4 + e) | 0;
    h5 = (h5 + f) | 0;
    h6 = (h6 + g) | 0;
    h7 = (h7 + h) | 0;
  }
  state[0] = h0;
  state[1] = h1;
  state[2] = h2;
  state[3] = h3;
  state[4] = h4;
  state[5] = h5;
  state[6] = h6;
  state[7] = h7;
}

const fastState = new Int32Array(8);
const fastTail = new Uint8Array(128);

/** `sha256:<hex>` of `fastState`. */
function fastStateText(): string {
  let out = 'sha256:';
  for (let index = 0; index < 8; index += 1) {
    const word = fastState[index]!;
    out +=
      hexOfByte[(word >>> 24) & 0xff]! +
      hexOfByte[(word >>> 16) & 0xff]! +
      hexOfByte[(word >>> 8) & 0xff]! +
      hexOfByte[word & 0xff]!;
  }
  return out;
}

/**
 * SHA-256 of the first `length` bytes of `bytes`: the whole blocks read in place, the last
 * partial block and the padding in a small scratch block. The same digest as the reference,
 * without copying the input or allocating per call.
 */
function sha256Prefix(bytes: Uint8Array, length: number): string {
  fastState.set(SHA256_INITIAL);
  const whole = length - (length % 64);
  compressBlocks(fastState, bytes, 0, whole);
  const rest = length - whole;
  const tail = fastTail;
  tail.fill(0);
  for (let index = 0; index < rest; index += 1) tail[index] = bytes[whole + index]!;
  tail[rest] = 0x80;
  const tailLength = rest + 9 <= 64 ? 64 : 128;
  // The bit length, big-endian in the last eight bytes. Exact for any length a Uint8Array has.
  const high = Math.floor(length / 0x20000000);
  const low = (length * 8) >>> 0;
  tail[tailLength - 8] = high >>> 24;
  tail[tailLength - 7] = (high >>> 16) & 0xff;
  tail[tailLength - 6] = (high >>> 8) & 0xff;
  tail[tailLength - 5] = high & 0xff;
  tail[tailLength - 4] = low >>> 24;
  tail[tailLength - 3] = (low >>> 16) & 0xff;
  tail[tailLength - 2] = (low >>> 8) & 0xff;
  tail[tailLength - 1] = low & 0xff;
  compressBlocks(fastState, tail, 0, tailLength);
  return fastStateText();
}

/** The reference's digest, computed without its per-call copies. Exported for tests. */
export const sha256FontBytesFast = (bytes: Uint8Array): string =>
  sha256Prefix(bytes, bytes.byteLength);

const textEncoder = new TextEncoder();
/** Reused UTF-8 scratch for {@link sha256Text}; texts larger than its ceiling encode afresh. */
let textScratch = new Uint8Array(4096);
const MAX_TEXT_SCRATCH_BYTES = 1 << 20;

type NativeSha256 = (bytes: Uint8Array) => string;

/**
 * A native SHA-256 when the host has one, or null.
 *
 * The pure-JavaScript block above is correct everywhere and 40 to 80 times slower than the
 * platform's implementation. Layout hashes a token for every node it keys, and font
 * admission hashes every face it trusts, so on a long document the pure path spends more
 * time hashing than shaping. Node exposes `crypto.createHash` synchronously through
 * `process.getBuiltinModule` (20.16+, 22.3+, the engine's own floor) with no import that a
 * browser bundler would have to resolve, and Bun exposes `Bun.CryptoHasher`. Browsers have
 * only the asynchronous `crypto.subtle`, so they keep the pure path. The native answer is
 * checked against a known vector once; a runtime whose hasher disagrees falls back.
 */
function detectNativeSha256(): NativeSha256 | null {
  const host = globalThis as {
    Bun?: {
      CryptoHasher?: new (algorithm: string) => {
        update(b: Uint8Array): { digest(e: 'hex'): string };
      };
    };
    process?: { getBuiltinModule?: (id: string) => unknown };
  };
  let candidate: NativeSha256 | null = null;
  try {
    const Hasher = host.Bun?.CryptoHasher;
    if (Hasher) {
      candidate = (bytes) => new Hasher('sha256').update(bytes).digest('hex');
    } else if (typeof host.process?.getBuiltinModule === 'function') {
      const crypto = host.process.getBuiltinModule('node:crypto') as
        | { createHash?: (a: string) => { update(b: Uint8Array): { digest(e: 'hex'): string } } }
        | undefined;
      const createHash = crypto?.createHash;
      if (typeof createHash === 'function')
        candidate = (bytes) => createHash('sha256').update(bytes).digest('hex');
    }
    if (!candidate) return null;
    // FIPS 180-4 test vector for "abc".
    const probe = candidate(new Uint8Array([0x61, 0x62, 0x63]));
    return probe === 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'
      ? candidate
      : null;
  } catch {
    return null;
  }
}

const nativeSha256 = detectNativeSha256();

/** Whether this runtime hashes natively. Exported for tests and diagnostics. */
export const sha256IsNative = nativeSha256 !== null;

/** Synchronous platform-neutral SHA-256 used before font bytes cross into a shaping implementation. */
export const sha256FontBytes: (bytes: Uint8Array) => string = nativeSha256
  ? (bytes) => `sha256:${nativeSha256(bytes)}`
  : sha256FontBytesFast;

/**
 * `sha256FontBytes` of the UTF-8 encoding of `text` (as `TextEncoder` encodes it, a lone
 * surrogate as U+FFFD), encoded into a reused buffer instead of a new one per call.
 */
export const sha256Text: (text: string) => string = nativeSha256
  ? (text) => `sha256:${nativeSha256(textEncoder.encode(text))}`
  : sha256TextFast;

/** {@link sha256Text} without a native hasher. Exported for tests. */
export function sha256TextFast(text: string): string {
  // UTF-8 never takes more than three bytes per UTF-16 code unit.
  const most = text.length * 3;
  if (most > MAX_TEXT_SCRATCH_BYTES) return sha256FontBytesFast(textEncoder.encode(text));
  if (textScratch.length < most) {
    let size = textScratch.length;
    while (size < most) size *= 2;
    textScratch = new Uint8Array(size);
  }
  const { written } = textEncoder.encodeInto(text, textScratch);
  return sha256Prefix(textScratch, written);
}

/** Lower-case hex of a digest, in the `sha256:` form every caller compares. */
const digestText = (digest: ArrayBuffer): string =>
  `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`;

/**
 * SHA-256 of `bytes` in the same `sha256:<hex>` form, from the browser's `crypto.subtle` where
 * there is no synchronous native hasher.
 *
 * The digest runs in the platform's own implementation instead of the pure path, which is what a
 * browser otherwise pays for every font face it admits. The bytes are read after an `await`, so
 * the caller must own them: pass a copy no one else can write, never a host's buffer. Without
 * `crypto.subtle` (an insecure origin) or when it refuses, the synchronous hash answers.
 */
export function sha256BytesAsync(bytes: Uint8Array): Promise<string> {
  const subtle = nativeSha256
    ? undefined
    : (globalThis as { crypto?: { subtle?: SubtleCrypto } }).crypto?.subtle;
  return subtle ? sha256WithSubtle(bytes, subtle) : Promise.resolve(sha256FontBytes(bytes));
}

/** The `crypto.subtle` half of {@link sha256BytesAsync}. Exported for tests. */
export async function sha256WithSubtle(bytes: Uint8Array, subtle: SubtleCrypto): Promise<string> {
  try {
    return digestText(await subtle.digest('SHA-256', bytes as Uint8Array<ArrayBuffer>));
  } catch {
    return sha256FontBytes(bytes);
  }
}
