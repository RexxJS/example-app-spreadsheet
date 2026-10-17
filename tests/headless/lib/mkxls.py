# Writes XLS (BIFF8) fixtures into tests/headless/lib/xls_fixtures.ts: the
# records written here by hand from [MS-XLS], put in a compound file
# ([MS-CFB]) built here too; base64. Not made by Excel or GridCraft.
import struct, base64
def rec(t, data): return struct.pack('<HH', t, len(data)) + data
def ustr8(s): return struct.pack('<BB', len(s), 0) + s.encode('latin1')
def ustr16(s): return struct.pack('<HB', len(s), 0) + s.encode('latin1')
BOF_G = rec(0x0809, struct.pack('<HHHHII', 0x0600, 0x0005, 0x0DBB, 0x07CC, 0, 0x06))
BOF_S = rec(0x0809, struct.pack('<HHHHII', 0x0600, 0x0010, 0x0DBB, 0x07CC, 0, 0x06))
EOF_ = rec(0x000A, b'')
strings = ['Item', 'pens', 'ink', '007', 'Total']
sst = struct.pack('<II', len(strings), len(strings))
for s in strings: sst += struct.pack('<HB', len(s), 0) + s.encode('latin1')
# the second string split across a CONTINUE (high bit: wide characters there)
sst_a = sst[:8 + 7 + 3]  # 'Item' whole, then 'pens': count + flags + 'p'?  computed below
# Build SST + CONTINUE: cut inside 'pens' after 2 characters.
cut = 8 + (3 + 4) + (3 + 2)
sst_rec = rec(0x00FC, sst[:cut]) + rec(0x003C, b'\x01' + 'ns'.encode('utf-16-le') + sst[cut + 2:])
fmt = rec(0x041E, struct.pack('<H', 164) + ustr16('yyyy-mm-dd'))
# XF records: 0..15 style xfs (plain), 15 cell default, 16 the date one, 17 the 0.00 one.
xfs = b''.join(rec(0x00E0, struct.pack('<HHHBBBBIIH', 0, 0, 0xFFF5 if i < 15 else 0, 0x20, 0, 0, 0, 0, 0, 0x20C0)) for i in range(16))
xfs += rec(0x00E0, struct.pack('<HHHBBBBIIH', 0, 164, 0, 0x20, 0, 0, 0, 0, 0, 0x20C0))
xfs += rec(0x00E0, struct.pack('<HHHBBBBIIH', 0, 2, 0, 0x20, 0, 0, 0, 0, 0, 0x20C0))
datemode = rec(0x0022, struct.pack('<H', 0))
supbook = rec(0x01AE, struct.pack('<HH', 2, 0x0401))
extern = rec(0x0017, struct.pack('<H', 2) + struct.pack('<HHH', 0, 0, 0) + struct.pack('<HHH', 0, 1, 1))
name = rec(0x0018, struct.pack('<HBBHHHBBBB', 0, 0, 3, 7, 0, 0, 0, 0, 0, 0) + b'\x00' + b'QTY' + struct.pack('<BHHHBB', 0x3B, 0, 1, 2, 1, 1)[:0] + bytes([0x3B]) + struct.pack('<HHHHH', 0, 1, 2, 1, 1))
def sheet_rec(name, pos, hidden=0): return rec(0x0085, struct.pack('<IBB', pos, hidden, 0) + ustr8(name))
def number(r, c, xf, v): return rec(0x0203, struct.pack('<HHHd', r, c, xf, v))
def rk_int(v): return (v << 2) | 2
def rkrec(r, c, xf, v): return rec(0x027E, struct.pack('<HHHI', r, c, xf, v))
def labelsst(r, c, xf, i): return rec(0x00FD, struct.pack('<HHHI', r, c, xf, i))
def boolerr(r, c, v, err): return rec(0x0205, struct.pack('<HHHBB', r, c, 15, v, err))
def formula(r, c, xf, result, rgce): return rec(0x0006, struct.pack('<HHH', r, c, xf) + result + struct.pack('<HI', 0, 0) + struct.pack('<H', len(rgce)) + rgce)
num_res = lambda v: struct.pack('<d', v)
str_res = struct.pack('<BBHBBH', 0, 0, 0, 0, 0, 0xFFFF)
ref = lambda row, col, rrel=1, crel=1: struct.pack('<HH', row, col | (crel << 14) | (rrel << 15))
area = lambda r1, r2, c1, c2: struct.pack('<HHHH', r1, r2, c1 | 0xC000, c2 | 0xC000)
s1 = BOF_S
s1 += labelsst(0, 0, 15, 0) + labelsst(1, 0, 15, 1) + labelsst(2, 0, 15, 2) + labelsst(3, 0, 15, 3)
s1 += number(1, 1, 17, 3.5) + rkrec(2, 1, 15, rk_int(12))
s1 += rec(0x00BD, struct.pack('<HH', 4, 1) + struct.pack('<HI', 15, rk_int(7)) + struct.pack('<HI', 15, (150 << 2) | 3) + struct.pack('<H', 2))  # B5 7, C5 1.5
s1 += boolerr(5, 0, 1, 0) + boolerr(5, 1, 7, 1)  # A6 TRUE, B6 #DIV/0!
s1 += number(6, 0, 16, 45306)  # A7 a date
# C2: =SUM(B2:B3)*2  ->  ptgArea, ptgAttr(sum), ptgInt 2, ptgMul
s1 += formula(1, 2, 15, num_res(31), bytes([0x25]) + area(1, 2, 1, 1) + bytes([0x19, 0x10, 0, 0]) + bytes([0x1E]) + struct.pack('<H', 2) + bytes([0x05]))
# C3: =IF(B3>10,"big","small")&"!"  ->  ref, int, gt, str, str, funcvar IF 3, str, concat
s1 += formula(2, 2, 15, str_res, bytes([0x24]) + ref(2, 1) + bytes([0x1E]) + struct.pack('<H', 10) + bytes([0x0D]) + bytes([0x17]) + ustr8('big') + bytes([0x17]) + ustr8('small') + bytes([0x42, 3]) + struct.pack('<H', 1) + bytes([0x17]) + ustr8('!') + bytes([0x08]))
s1 += rec(0x0207, ustr16('big!'))
# C4: =Other!A1+$B$2  ->  ref3d (ixti 1), absolute ref, add
s1 += formula(3, 2, 15, num_res(13.5), bytes([0x3A]) + struct.pack('<H', 1) + ref(0, 0) + bytes([0x24]) + ref(1, 1, 0, 0) + bytes([0x03]))
# C5: =ROUND(-B2,0)  -> ref, uminus, int 0, func ROUND (27)
s1 += formula(4, 2, 15, num_res(-4), bytes([0x24]) + ref(1, 1) + bytes([0x13]) + bytes([0x1E]) + struct.pack('<H', 0) + bytes([0x41]) + struct.pack('<H', 27))
# C6: an add-in function (255): its cached value kept
s1 += formula(5, 2, 15, num_res(99), bytes([0x17]) + ustr8('X') + bytes([0x42, 1]) + struct.pack('<H', 255))
s1 += rec(0x00E5, struct.pack('<H', 1) + struct.pack('<HHHH', 7, 7, 0, 2))  # A8:C8
s1 += labelsst(7, 0, 15, 4)
s1 += number(300, 0, 15, 1)  # past the sheet's edge
s1 += EOF_
s2 = BOF_S + number(0, 0, 15, 10) + EOF_
def globals(p1, p2):
    return (BOF_G + datemode + fmt + xfs + supbook + extern + sheet_rec('Sales', p1) + sheet_rec('Other', p2, 1) + sst_rec + EOF_)
g = globals(0, 0)
p1 = len(g); p2 = p1 + len(s1)
g = globals(p1, p2)
assert len(g) == p1
workbook = g + s1 + s2
def cfb(streams):
    # One FAT sector, one directory sector; streams padded to whole sectors
    # (each at least 4096 bytes, so in the FAT, not the mini stream).
    SECT = 512
    data = b''; starts = []
    fat = []
    sect = 2  # 0 FAT, 1 directory
    for name, body in streams:
        body = body + b'\x00' * (max(4096, -(-len(body) // SECT) * SECT) - len(body))
        n = len(body) // SECT
        starts.append((name, sect, len(body)))
        data += body
        sect += n
    fat = [0xFFFFFFFD, 0xFFFFFFFE]
    for name, start, size in starts:
        n = size // SECT
        fat += [start + k + 1 for k in range(n - 1)] + [0xFFFFFFFE]
    assert len(fat) <= 128
    fat += [0xFFFFFFFF] * (128 - len(fat))
    def entry(name, typ, start, size, child=0xFFFFFFFF, left=0xFFFFFFFF, right=0xFFFFFFFF):
        nm = (name.encode('utf-16-le') + b'\x00\x00')
        return nm + b'\x00' * (64 - len(nm)) + struct.pack('<HBBIII', len(nm), typ, 1, left, right, child) + b'\x00' * 16 + struct.pack('<I', 0) + b'\x00' * 16 + struct.pack('<IQ', start, size)
    d = entry('Root Entry', 5, 0xFFFFFFFE, 0, child=1)
    for i, (name, start, size) in enumerate(starts):
        d += entry(name, 2, start, size, right=(i + 2) if i + 1 < len(starts) else 0xFFFFFFFF)
    d += b'\x00' * (SECT - len(d))
    hdr = bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) + b'\x00' * 16 + struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\x00' * 6 + struct.pack('<IIIIIIIII', 0, 1, 1, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0)
    hdr = hdr[:0x4C] + struct.pack('<I', 0) + struct.pack('<I', 0xFFFFFFFF) * 108
    assert len(hdr) == 512, len(hdr)
    return hdr + struct.pack('<128I', *fat) + d + data
book = cfb([('Workbook', workbook)])
open('/Users/paul/scm/spraedsheet/tests/headless/lib/xls_fixtures.ts', 'w').write(f'''// An XLS (BIFF8) file made by hand for the tests: its records and its
// compound file written in mkxls.py beside this file, from [MS-XLS] and
// [MS-CFB] (not by Excel or GridCraft), as base64. BOOK: two sheets (the
// second hidden), shared strings (one broken across a CONTINUE record,
// wide there), a date format, NUMBER, RK, MULRK, LABELSST, BOOLERR cells,
// formulas (SUM by ptgAttr, IF with a string result and its STRING
// record, a reference to the other sheet, an absolute one, a unary minus
// and ROUND, an add-in function kept as its value), a merge, a cell past
// the sheet's edge.
export const BOOK = "{base64.b64encode(book).decode()}";
''')
print(len(book))
