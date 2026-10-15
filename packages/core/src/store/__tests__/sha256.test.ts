import { describe, expect, test } from 'bun:test';
import {
  sha256BytesAsync,
  sha256FontBytes,
  sha256FontBytesPure,
  sha256IsNative,
  sha256WithSubtle,
} from '../package/sha256.ts';

describe('sha256FontBytes', () => {
  test('matches the FIPS 180-4 vectors', () => {
    const text = (s: string) => new TextEncoder().encode(s);
    expect(sha256FontBytesPure(text(''))).toBe(
      'sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    );
    expect(sha256FontBytesPure(text('abc'))).toBe(
      'sha256:ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'
    );
    expect(sha256FontBytes(text('abc'))).toBe(sha256FontBytesPure(text('abc')));
  });

  test('the native path, when the host has one, agrees with the pure path byte for byte', () => {
    // Bun always has a native hasher, so this exercises the fast path in the test suite.
    expect(sha256IsNative).toBe(true);
    for (const length of [0, 1, 55, 56, 63, 64, 65, 1000, 4096, 70_000]) {
      const bytes = new Uint8Array(length);
      for (let i = 0; i < length; i++) bytes[i] = (i * 7919 + length) & 0xff;
      expect(sha256FontBytes(bytes)).toBe(sha256FontBytesPure(bytes));
    }
    // A subarray view hashes the view, not the whole buffer beneath it.
    const buffer = new Uint8Array(200).fill(9);
    const view = buffer.subarray(50, 120);
    expect(sha256FontBytes(view)).toBe(sha256FontBytesPure(view));
    expect(sha256FontBytes(view)).not.toBe(sha256FontBytes(buffer));
  });
});

describe('sha256BytesAsync', () => {
  const sample = (length: number) => {
    const bytes = new Uint8Array(length);
    for (let i = 0; i < length; i++) bytes[i] = (i * 131 + 7) & 0xff;
    return bytes;
  };

  test('the platform digest a browser uses agrees with the pure path byte for byte', async () => {
    const subtle = globalThis.crypto.subtle;
    for (const length of [0, 3, 64, 1000, 70_000]) {
      const bytes = sample(length);
      expect(await sha256WithSubtle(bytes, subtle)).toBe(sha256FontBytesPure(bytes));
    }
    // A subarray view digests the view, not the buffer beneath it.
    const view = sample(200).subarray(50, 120);
    expect(await sha256WithSubtle(view, subtle)).toBe(sha256FontBytesPure(view));
  });

  test('a digest the platform refuses falls back to the synchronous hash', async () => {
    const refusing = {
      digest: () => Promise.reject(new Error('not allowed')),
    } as unknown as SubtleCrypto;
    const bytes = sample(100);
    expect(await sha256WithSubtle(bytes, refusing)).toBe(sha256FontBytes(bytes));
  });

  test('answers what the synchronous hash answers', async () => {
    const bytes = sample(5000);
    expect(await sha256BytesAsync(bytes)).toBe(sha256FontBytes(bytes));
  });
});
