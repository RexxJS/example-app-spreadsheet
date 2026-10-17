// The ciphers an encrypted Office file is opened with: SHA-1 and SHA-512
// (FIPS 180-4; SHA-512's 64-bit words as two 32-bit halves, the page's
// engine having no BigInt) and AES decryption (FIPS 197: 128, 192 or
// 256-bit keys), in ECB and CBC. Pure; only what opening a file needs.

// ---- SHA-1.
export function sha1(m: Uint8Array): Uint8Array {
  const len = m.length, blocks = ((len + 8) >> 6) + 1, w = new Int32Array(blocks * 16);
  for (let i = 0; i < len; i++) w[i >> 2] |= m[i] << (24 - (i % 4) * 8);
  w[len >> 2] |= 0x80 << (24 - (len % 4) * 8);
  w[blocks * 16 - 1] = len * 8;
  const h = [0x67452301, 0xefcdab89 | 0, 0x98badcfe | 0, 0x10325476, 0xc3d2e1f0 | 0], x = new Int32Array(80);
  for (let i = 0; i < w.length; i += 16) {
    let [a, b, c, d, e] = h;
    for (let t = 0; t < 80; t++) {
      if (t < 16) x[t] = w[i + t];
      else { const v = x[t - 3] ^ x[t - 8] ^ x[t - 14] ^ x[t - 16]; x[t] = (v << 1) | (v >>> 31); }
      const f = t < 20 ? (b & c) | (~b & d) : t < 40 ? b ^ c ^ d : t < 60 ? (b & c) | (b & d) | (c & d) : b ^ c ^ d;
      const k = t < 20 ? 0x5a827999 : t < 40 ? 0x6ed9eba1 : t < 60 ? 0x8f1bbcdc | 0 : 0xca62c1d6 | 0;
      const tmp = (((a << 5) | (a >>> 27)) + f + e + k + x[t]) | 0;
      e = d; d = c; c = (b << 30) | (b >>> 2); b = a; a = tmp;
    }
    h[0] = (h[0] + a) | 0; h[1] = (h[1] + b) | 0; h[2] = (h[2] + c) | 0; h[3] = (h[3] + d) | 0; h[4] = (h[4] + e) | 0;
  }
  return bytesOf(h);
}

function bytesOf(ws: readonly number[]): Uint8Array {
  const out = new Uint8Array(ws.length * 4);
  ws.forEach((v, i) => { out[i * 4] = v >>> 24; out[i * 4 + 1] = (v >>> 16) & 255; out[i * 4 + 2] = (v >>> 8) & 255; out[i * 4 + 3] = v & 255; });
  return out;
}

// ---- SHA-512, its words as [high, low] 32-bit pairs.
const K512 = [
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc, 0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019, 0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118,
  0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2, 0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694,
  0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3, 0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65, 0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
  0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4, 0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725, 0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70,
  0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df, 0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b,
  0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001, 0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30, 0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
  0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8, 0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb, 0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3,
  0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec, 0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b,
  0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207, 0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178, 0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
  0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c, 0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a, 0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
];
const H512 = [0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1, 0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179];

export function sha512(m: Uint8Array): Uint8Array {
  const len = m.length, blocks = ((len + 16) >> 7) + 1, w = new Int32Array(blocks * 32);
  for (let i = 0; i < len; i++) w[i >> 2] |= m[i] << (24 - (i % 4) * 8);
  w[len >> 2] |= 0x80 << (24 - (len % 4) * 8);
  w[blocks * 32 - 1] = len * 8;
  w[blocks * 32 - 2] = Math.floor(len / 0x20000000);
  const h = Int32Array.from(H512.map((x) => x | 0)), W = new Int32Array(160);
  // 64-bit sums of [hi, lo] pairs, carried by hand.
  for (let i = 0; i < w.length; i += 32) {
    for (let t = 0; t < 80; t++) {
      if (t < 16) { W[t * 2] = w[i + t * 2]; W[t * 2 + 1] = w[i + t * 2 + 1]; continue; }
      const [xh, xl] = [W[(t - 15) * 2], W[(t - 15) * 2 + 1]], [yh, yl] = [W[(t - 2) * 2], W[(t - 2) * 2 + 1]];
      // σ0 = rotr1 ^ rotr8 ^ shr7; σ1 = rotr19 ^ rotr61 ^ shr6
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      const s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
      const s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
      let lo = (s0l >>> 0) + (s1l >>> 0) + (W[(t - 7) * 2 + 1] >>> 0) + (W[(t - 16) * 2 + 1] >>> 0);
      const hi = (s0h + s1h + W[(t - 7) * 2] + W[(t - 16) * 2] + Math.floor(lo / 0x100000000)) | 0;
      W[t * 2] = hi; W[t * 2 + 1] = lo | 0;
    }
    let [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl] = h;
    for (let t = 0; t < 80; t++) {
      // Σ1(e) = rotr14 ^ rotr18 ^ rotr41; ch = (e & f) ^ (~e & g)
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
      let lo = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[t * 2 + 1] >>> 0) + (W[t * 2 + 1] >>> 0);
      const t1h = (hh + S1h + chh + K512[t * 2] + W[t * 2] + Math.floor(lo / 0x100000000)) | 0, t1l = lo | 0;
      // Σ0(a) = rotr28 ^ rotr34 ^ rotr39; maj
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      lo = (S0l >>> 0) + (mjl >>> 0);
      const t2h = (S0h + mjh + Math.floor(lo / 0x100000000)) | 0, t2l = lo | 0;
      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      lo = (dl >>> 0) + (t1l >>> 0);
      eh = (dh + t1h + Math.floor(lo / 0x100000000)) | 0; el = lo | 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      lo = (t1l >>> 0) + (t2l >>> 0);
      ah = (t1h + t2h + Math.floor(lo / 0x100000000)) | 0; al = lo | 0;
    }
    const v = [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl];
    for (let k = 0; k < 16; k += 2) {
      const lo = (h[k + 1] >>> 0) + (v[k + 1] >>> 0);
      h[k] = (h[k] + v[k] + Math.floor(lo / 0x100000000)) | 0;
      h[k + 1] = lo | 0;
    }
  }
  return bytesOf(Array.from(h));
}

// ---- AES decryption.
const SBOX = new Uint8Array(256), INV = new Uint8Array(256);
(() => {
  // The S-box from GF(2^8) inverses and the affine map.
  let p = 1, q = 1;
  do {
    p = p ^ (p << 1) ^ (p & 0x80 ? 0x1b : 0); p &= 255;
    q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 255; if (q & 0x80) q ^= 0x09;
    const x = q ^ ((q << 1) | (q >> 7)) ^ ((q << 2) | (q >> 6)) ^ ((q << 3) | (q >> 5)) ^ ((q << 4) | (q >> 4));
    SBOX[p] = (x ^ 0x63) & 255;
  } while (p !== 1);
  SBOX[0] = 0x63;
  for (let i = 0; i < 256; i++) INV[SBOX[i]] = i;
})();
const xt = (b: number) => ((b << 1) ^ (b & 0x80 ? 0x1b : 0)) & 255;
const mul = (a: number, b: number) => { let r = 0; for (; b; b >>= 1, a = xt(a)) if (b & 1) r ^= a; return r; };

// The round keys of a 16, 24 or 32-byte key.
export function expandKey(key: Uint8Array): Uint8Array {
  const nk = key.length / 4, rounds = nk + 6, out = new Uint8Array(16 * (rounds + 1));
  if (![4, 6, 8].includes(nk)) throw new RangeError(`an AES key of ${key.length} bytes`);
  out.set(key);
  let rcon = 1;
  for (let i = nk; i < 4 * (rounds + 1); i++) {
    let t = out.slice((i - 1) * 4, i * 4);
    if (i % nk === 0) { t = new Uint8Array([SBOX[t[1]] ^ rcon, SBOX[t[2]], SBOX[t[3]], SBOX[t[0]]]); rcon = xt(rcon); }
    else if (nk > 6 && i % nk === 4) t = t.map((b) => SBOX[b]);
    for (let k = 0; k < 4; k++) out[i * 4 + k] = out[(i - nk) * 4 + k] ^ t[k];
  }
  return out;
}

// One 16-byte block decrypted with the round keys.
export function decryptBlock(rk: Uint8Array, block: Uint8Array): Uint8Array {
  const rounds = rk.length / 16 - 1, s = new Uint8Array(16);
  for (let i = 0; i < 16; i++) s[i] = block[i] ^ rk[rounds * 16 + i];
  for (let r = rounds - 1; r >= 0; r--) {
    // Inverse shift rows, inverse sub bytes.
    const t = new Uint8Array(16);
    for (let c = 0; c < 4; c++) for (let row = 0; row < 4; row++) t[((c + row) % 4) * 4 + row] = INV[s[c * 4 + row]];
    for (let i = 0; i < 16; i++) t[i] ^= rk[r * 16 + i];
    if (r > 0) for (let c = 0; c < 4; c++) {
      const [a, b, cc, d] = [t[c * 4], t[c * 4 + 1], t[c * 4 + 2], t[c * 4 + 3]];
      s[c * 4] = mul(a, 14) ^ mul(b, 11) ^ mul(cc, 13) ^ mul(d, 9);
      s[c * 4 + 1] = mul(a, 9) ^ mul(b, 14) ^ mul(cc, 11) ^ mul(d, 13);
      s[c * 4 + 2] = mul(a, 13) ^ mul(b, 9) ^ mul(cc, 14) ^ mul(d, 11);
      s[c * 4 + 3] = mul(a, 11) ^ mul(b, 13) ^ mul(cc, 9) ^ mul(d, 14);
    } else s.set(t);
  }
  return s;
}

export function decryptEcb(key: Uint8Array, data: Uint8Array): Uint8Array {
  const rk = expandKey(key), out = new Uint8Array(data.length - (data.length % 16));
  for (let i = 0; i < out.length; i += 16) out.set(decryptBlock(rk, data.subarray(i, i + 16)), i);
  return out;
}

export function decryptCbc(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Uint8Array {
  const rk = expandKey(key), out = new Uint8Array(data.length - (data.length % 16));
  let prev = iv.subarray(0, 16);
  for (let i = 0; i < out.length; i += 16) {
    const block = data.subarray(i, i + 16), plain = decryptBlock(rk, block);
    for (let k = 0; k < 16; k++) out[i + k] = plain[k] ^ prev[k];
    prev = block;
  }
  return out;
}

export const hex = (b: Uint8Array): string => Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
