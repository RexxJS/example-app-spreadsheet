// An encrypted Office file ([MS-OFFCRYPTO]): a compound file holding
// EncryptionInfo and EncryptedPackage, the package an XLSX once its
// password opens it. Agile encryption (version 4.4: its XML says the
// cipher, the hash and the spin count; AES with SHA-512, or SHA-1, as the
// file says) and Standard encryption (3.2 or 4.2: AES-128 with SHA-1, the
// key from Standard's own derivation), each password checked by the
// file's verifier first. Pure.
import { compoundStreams } from "./xls.ts";
import { decryptCbc, decryptEcb, sha1, sha512 } from "./crypto.ts";
import { attr, all, kid, parseXml } from "./xml.ts";
import { fromBase64 } from "./csv.ts";
import { utf8Text } from "./zip.ts";

const u16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8);
const u32 = (b: Uint8Array, i: number) => (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0;
const join = (...xs: Uint8Array[]): Uint8Array => { const out = new Uint8Array(xs.reduce((n, x) => n + x.length, 0)); let o = 0; for (const x of xs) { out.set(x, o); o += x.length; } return out; };
const le32 = (n: number) => new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
const utf16 = (s: string) => { const out = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 255; out[i * 2 + 1] = s.charCodeAt(i) >> 8; } return out; };
const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((x, i) => x === b[i]);
// A hash's output cut or padded (with 0x36) to n bytes.
const fit = (h: Uint8Array, n: number) => (h.length >= n ? h.slice(0, n) : join(h, new Uint8Array(n - h.length).fill(0x36)));

export class WrongPassword extends Error {}

// Whether these bytes are an encrypted Office file.
export function isEncrypted(b: Uint8Array): boolean {
  if (b.length < 512 || b[0] !== 0xd0 || b[1] !== 0xcf) return false;
  try {
    return compoundStreams(b).has("EncryptionInfo");
  } catch {
    return false;
  }
}

// The package (an XLSX's bytes) of an encrypted file, by its password.
export function decryptPackage(b: Uint8Array, password: string): Uint8Array {
  const s = compoundStreams(b), info = s.get("EncryptionInfo"), pkg = s.get("EncryptedPackage");
  if (!info || !pkg) throw new RangeError("not an encrypted Office file");
  const major = u16(info, 0), minor = u16(info, 2), size = u32(pkg, 0) + u32(pkg, 4) * 0x100000000, data = pkg.subarray(8);
  if (major === 4 && minor === 4) return agile(info, data, size, password);
  if ((major === 3 || major === 4) && minor === 2) return standard(info, data, size, password);
  throw new RangeError(`encryption version ${major}.${minor} is not one the app opens (Agile 4.4, Standard 3.2 or 4.2)`);
}

function agile(info: Uint8Array, data: Uint8Array, size: number, password: string): Uint8Array {
  const xml = parseXml(utf8Text(info.subarray(8)));
  const kd = kid(xml, "keyData"), ke = all(xml, "encryptedKey")[0];
  if (!kd || !ke) throw new RangeError("an Agile file with no key data");
  const hashOf = (name: string | undefined) => (name === "SHA512" ? sha512 : name === "SHA1" ? sha1 : (() => { throw new RangeError(`the hash ${name} is not one the app has (SHA512, SHA1)`); })());
  if ((attr(ke, "cipherAlgorithm") ?? "AES") !== "AES" || (attr(ke, "cipherChaining") ?? "ChainingModeCBC") !== "ChainingModeCBC") throw new RangeError("a cipher other than AES-CBC");
  const H = hashOf(attr(ke, "hashAlgorithm")), spins = Number(attr(ke, "spinCount") ?? 100000), keyBytes = Number(attr(ke, "keyBits") ?? 256) / 8;
  const salt = fromBase64(attr(ke, "saltValue") ?? ""), blockSize = Number(attr(ke, "blockSize") ?? 16);
  let h = H(join(salt, utf16(password)));
  for (let i = 0; i < spins; i++) h = H(join(le32(i), h));
  const keyFor = (block: number[]) => fit(H(join(h, new Uint8Array(block))), keyBytes);
  const iv = fit(salt, blockSize);
  const input = decryptCbc(keyFor([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]), iv, fromBase64(attr(ke, "encryptedVerifierHashInput") ?? ""));
  const value = decryptCbc(keyFor([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]), iv, fromBase64(attr(ke, "encryptedVerifierHashValue") ?? ""));
  const hashSize = H(new Uint8Array(0)).length;
  if (!same(H(input.slice(0, Number(attr(ke, "saltSize") ?? 16))), value.slice(0, hashSize))) throw new WrongPassword("the password is wrong");
  const secret = decryptCbc(keyFor([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]), iv, fromBase64(attr(ke, "encryptedKeyValue") ?? "")).slice(0, keyBytes);
  // The package in 4096-byte segments, each its own IV.
  const KH = hashOf(attr(kd, "hashAlgorithm")), kSalt = fromBase64(attr(kd, "saltValue") ?? ""), kBlock = Number(attr(kd, "blockSize") ?? 16);
  const out = new Uint8Array(Math.ceil(data.length / 16) * 16);
  for (let seg = 0, o = 0; o < data.length; seg++, o += 4096) {
    const piece = data.subarray(o, Math.min(data.length, o + 4096));
    out.set(decryptCbc(secret, fit(KH(join(kSalt, le32(seg))), kBlock), piece), o);
  }
  return out.slice(0, size);
}

function standard(info: Uint8Array, data: Uint8Array, size: number, password: string): Uint8Array {
  const headerSize = u32(info, 8), header = info.subarray(12, 12 + headerSize), verifier = info.subarray(12 + headerSize);
  const algId = u32(header, 8), keyBits = u32(header, 16);
  if (algId !== 0x660e && algId !== 0x660f && algId !== 0x6610) throw new RangeError("a Standard file not encrypted with AES (RC4 is not opened)");
  const saltSize = u32(verifier, 0), salt = verifier.subarray(4, 4 + saltSize);
  const encVerifier = verifier.subarray(4 + saltSize, 20 + saltSize), encHash = verifier.subarray(24 + saltSize, 56 + saltSize);
  let h = sha1(join(salt, utf16(password)));
  for (let i = 0; i < 50000; i++) h = sha1(join(le32(i), h));
  h = sha1(join(h, le32(0)));
  const pad = (c: number) => { const b = new Uint8Array(64).fill(c); for (let i = 0; i < h.length; i++) b[i] ^= h[i]; return sha1(b); };
  const key = join(pad(0x36), pad(0x5c)).slice(0, keyBits / 8);
  const v = decryptEcb(key, encVerifier), vh = decryptEcb(key, encHash);
  if (!same(sha1(v), vh.slice(0, 20))) throw new WrongPassword("the password is wrong");
  return decryptEcb(key, data).slice(0, size);
}
