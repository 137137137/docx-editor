import { describe, expect, test } from 'bun:test';
import {
  sha256BytesAsync,
  sha256FontBytes,
  sha256FontBytesFast,
  sha256FontBytesPure,
  sha256IsNative,
  sha256Text,
  sha256TextFast,
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

describe('the browser path', () => {
  test('agrees with the reference at every padding boundary and on a view', () => {
    for (const length of [0, 1, 3, 55, 56, 57, 63, 64, 65, 119, 120, 127, 128, 129, 1000, 70_000]) {
      const bytes = new Uint8Array(length);
      for (let i = 0; i < length; i++) bytes[i] = (i * 7919 + length) & 0xff;
      expect(sha256FontBytesFast(bytes)).toBe(sha256FontBytesPure(bytes));
    }
    const view = new Uint8Array(300).fill(5).subarray(17, 211);
    expect(sha256FontBytesFast(view)).toBe(sha256FontBytesPure(view));
  });

  test('a text digest is the digest of its UTF-8 bytes, whatever the text holds', () => {
    const encode = (text: string) => new TextEncoder().encode(text);
    const texts = [
      '',
      'abc',
      'x'.repeat(55),
      'x'.repeat(56),
      'é'.repeat(40),
      '\u4e2d\u6587'.repeat(500),
      '\ud83d\ude00 emoji',
      'lone \ud800 surrogate',
      '\0framed\0',
      'y'.repeat(5000),
      'z'.repeat(400_000),
    ];
    for (const text of texts) {
      const expected = sha256FontBytesPure(encode(text));
      expect(sha256TextFast(text)).toBe(expected);
      expect(sha256Text(text)).toBe(expected);
    }
    // The scratch buffer is reused: a short text after a long one hashes only its own bytes.
    expect(sha256TextFast('abc')).toBe(sha256FontBytesPure(encode('abc')));
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
