// ZIP archives, as XLSX and ODS files are: read (the central directory's
// entries, each stored or deflated) and written (stored or deflated), with
// DEFLATE (RFC 1951) both ways and CRC-32. Pure: bytes in, bytes out.

// ---- CRC-32 (the ZIP polynomial, reflected).
const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();
export function crc32(b: Uint8Array): number {
  let c = -1;
  for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 255] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

// ---- Inflate.
const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

// A Huffman code from its lengths: the count of each length, and the
// symbols in code order.
interface Huff { readonly counts: Int32Array; readonly symbols: Int32Array }
function huff(lengths: ArrayLike<number>): Huff {
  const counts = new Int32Array(16), offs = new Int32Array(16), symbols = new Int32Array(lengths.length);
  for (let i = 0; i < lengths.length; i++) counts[lengths[i]]++;
  counts[0] = 0;
  for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
  for (let i = 0; i < lengths.length; i++) if (lengths[i]) symbols[offs[lengths[i]]++] = i;
  return { counts, symbols };
}

export function inflate(src: Uint8Array): Uint8Array {
  let pos = 0, bit = 0, out = new Uint8Array(Math.max(1024, src.length * 4)), n = 0;
  const need = (k: number) => {
    if (n + k <= out.length) return;
    let size = out.length * 2;
    while (size < n + k) size *= 2;
    const bigger = new Uint8Array(size);
    bigger.set(out.subarray(0, n));
    out = bigger;
  };
  const bits = (k: number): number => {
    let v = 0;
    for (let i = 0; i < k; i++) {
      if (pos >= src.length) throw new RangeError("a deflated stream that ends too soon");
      v |= ((src[pos] >> bit) & 1) << i;
      if (++bit === 8) { bit = 0; pos++; }
    }
    return v;
  };
  const decode = (h: Huff): number => {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = h.counts[len];
      if (code - count < first) return h.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new RangeError("a bad Huffman code in a deflated stream");
  };
  const fixedLit = huff(Array.from({ length: 288 }, (_, i) => (i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8)));
  const fixedDist = huff(Array.from({ length: 30 }, () => 5));
  for (let last = 0; !last;) {
    last = bits(1);
    const type = bits(2);
    if (type === 0) {
      if (bit) { bit = 0; pos++; }
      const len = src[pos] | (src[pos + 1] << 8);
      pos += 4;
      if (pos + len > src.length) throw new RangeError("a stored block that ends too soon");
      need(len);
      out.set(src.subarray(pos, pos + len), n);
      n += len;
      pos += len;
      continue;
    }
    let lit = fixedLit, dist = fixedDist;
    if (type === 2) {
      const hlit = bits(5) + 257, hdist = bits(5) + 1, hclen = bits(4) + 4;
      const cl = new Array(19).fill(0);
      for (let i = 0; i < hclen; i++) cl[ORDER[i]] = bits(3);
      const clh = huff(cl), lens: number[] = [];
      while (lens.length < hlit + hdist) {
        const sym = decode(clh);
        if (sym < 16) lens.push(sym);
        else if (sym === 16) { const prev = lens[lens.length - 1]; for (let r = 3 + bits(2); r > 0; r--) lens.push(prev); }
        else if (sym === 17) for (let r = 3 + bits(3); r > 0; r--) lens.push(0);
        else for (let r = 11 + bits(7); r > 0; r--) lens.push(0);
      }
      lit = huff(lens.slice(0, hlit));
      dist = huff(lens.slice(hlit));
    } else if (type !== 1) throw new RangeError("a deflated block of an unknown type");
    for (;;) {
      const sym = decode(lit);
      if (sym < 256) { need(1); out[n++] = sym; continue; }
      if (sym === 256) break;
      const li = sym - 257, len = LEN_BASE[li] + bits(LEN_EXTRA[li]);
      const di = decode(dist), d = DIST_BASE[di] + bits(DIST_EXTRA[di]);
      if (d > n) throw new RangeError("a deflated stream reaching back before its start");
      need(len);
      for (let i = 0; i < len; i++, n++) out[n] = out[n - d];
    }
  }
  return out.slice(0, n);
}

// ---- Deflate: fixed Huffman codes over LZ77 matches found by a hash of
// three bytes (smaller than stored, simple; not zlib's best).
export function deflate(src: Uint8Array): Uint8Array {
  const out: number[] = [];
  let acc = 0, nbits = 0;
  const put = (v: number, k: number) => {
    for (let i = 0; i < k; i++) {
      acc |= ((v >> i) & 1) << nbits;
      if (++nbits === 8) { out.push(acc); acc = 0; nbits = 0; }
    }
  };
  // A Huffman code goes in most significant bit first.
  const code = (v: number, k: number) => { for (let i = k - 1; i >= 0; i--) put((v >> i) & 1, 1); };
  const literal = (c: number) => {
    if (c < 144) code(0x30 + c, 8); else if (c < 256) code(0x190 + c - 144, 9); else if (c < 280) code(c - 256, 7); else code(0xc0 + c - 280, 8);
  };
  put(1, 1); put(1, 2);
  const head = new Map<number, number>(), prev = new Int32Array(src.length).fill(-1);
  for (let i = 0; i < src.length;) {
    let best = 0, at = 0;
    if (i + 2 < src.length) {
      const h = (src[i] << 16) | (src[i + 1] << 8) | src[i + 2];
      for (let j = head.get(h) ?? -1, tries = 0; j >= 0 && i - j <= 32768 && tries < 64; j = prev[j], tries++) {
        let l = 0;
        while (l < 258 && i + l < src.length && src[j + l] === src[i + l]) l++;
        if (l > best) { best = l; at = i - j; }
      }
      prev[i] = head.get(h) ?? -1;
      head.set(h, i);
    }
    if (best >= 3) {
      let li = 0;
      while (li < 28 && LEN_BASE[li + 1] <= best) li++;
      literal(257 + li);
      put(best - LEN_BASE[li], LEN_EXTRA[li]);
      let di = 0;
      while (di < 29 && DIST_BASE[di + 1] <= at) di++;
      code(di, 5);
      put(at - DIST_BASE[di], DIST_EXTRA[di]);
      for (let k = 1; k < best && i + k + 2 < src.length; k++) {
        const h = (src[i + k] << 16) | (src[i + k + 1] << 8) | src[i + k + 2];
        prev[i + k] = head.get(h) ?? -1;
        head.set(h, i + k);
      }
      i += best;
    } else literal(src[i++]);
  }
  literal(256);
  if (nbits) out.push(acc);
  return new Uint8Array(out);
}

// ---- UTF-8.
export function utf8Bytes(s: string): Uint8Array {
  const out: number[] = [];
  for (const ch of s) {
    const c = ch.codePointAt(0)!;
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return new Uint8Array(out);
}
export function utf8Text(b: Uint8Array): string {
  let s = "";
  for (let i = 0; i < b.length;) {
    const c = b[i];
    const n = c < 0x80 ? 1 : c < 0xe0 ? 2 : c < 0xf0 ? 3 : 4;
    let cp = n === 1 ? c : c & (0xff >> (n + 1));
    for (let k = 1; k < n; k++) cp = (cp << 6) | (b[i + k] & 63);
    s += String.fromCodePoint(cp);
    i += n;
  }
  return s;
}

// ---- The archive.
const u16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8);
const u32 = (b: Uint8Array, i: number) => (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0;

// Its entries by name, their bytes; refused when it is not a ZIP, or an
// entry is packed some way other than stored or deflated (or encrypted).
export function unzip(b: Uint8Array): Map<string, Uint8Array> {
  let end = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (u32(b, i) === 0x06054b50) { end = i; break; }
  if (end < 0) throw new RangeError("not a ZIP archive (an XLSX or ODS file is one)");
  const count = u16(b, end + 10);
  let at = u32(b, end + 16);
  const out = new Map<string, Uint8Array>();
  for (let e = 0; e < count; e++) {
    if (u32(b, at) !== 0x02014b50) throw new RangeError("a ZIP archive's directory is broken");
    const flags = u16(b, at + 8), method = u16(b, at + 10), size = u32(b, at + 20), nameLen = u16(b, at + 28), extra = u16(b, at + 30), comment = u16(b, at + 32), local = u32(b, at + 42);
    const name = utf8Text(b.subarray(at + 46, at + 46 + nameLen));
    at += 46 + nameLen + extra + comment;
    if (flags & 1) throw new RangeError(`${name} is encrypted in the archive`);
    const data = local + 30 + u16(b, local + 26) + u16(b, local + 28), raw = b.subarray(data, data + size);
    if (method === 0) out.set(name, raw.slice());
    else if (method === 8) out.set(name, inflate(raw));
    else throw new RangeError(`${name} is packed by method ${method}, not stored or deflated`);
  }
  return out;
}

// Entries as an archive, each deflated (stored when that is no smaller).
export function zip(entries: readonly (readonly [string, Uint8Array])[]): Uint8Array {
  const parts: Uint8Array[] = [], dir: number[] = [];
  let offset = 0;
  const le = (v: number, k: number) => Array.from({ length: k }, (_, i) => (v >>> (8 * i)) & 255);
  for (const [name, data] of entries) {
    const nm = utf8Bytes(name), packed = deflate(data), deflated = packed.length < data.length, body = deflated ? packed : data, crc = crc32(data);
    const head = [...le(0x04034b50, 4), 20, 0, ...le(0x0800, 2), ...le(deflated ? 8 : 0, 2), 0, 0, 0x21, 0, ...le(crc, 4), ...le(body.length, 4), ...le(data.length, 4), ...le(nm.length, 2), 0, 0];
    parts.push(new Uint8Array(head), nm, body);
    dir.push(...le(0x02014b50, 4), 20, 0, 20, 0, ...le(0x0800, 2), ...le(deflated ? 8 : 0, 2), 0, 0, 0x21, 0, ...le(crc, 4), ...le(body.length, 4), ...le(data.length, 4), ...le(nm.length, 2),
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...le(offset, 4), ...nm);
    offset += head.length + nm.length + body.length;
  }
  const end = [...le(0x06054b50, 4), 0, 0, 0, 0, ...le(entries.length, 2), ...le(entries.length, 2), ...le(dir.length, 4), ...le(offset, 4), 0, 0];
  const total = offset + dir.length + end.length, out = new Uint8Array(total);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  out.set(dir, o);
  out.set(end, o + dir.length);
  return out;
}
