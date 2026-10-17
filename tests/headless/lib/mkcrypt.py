# Writes tests/headless/lib/crypt_fixtures.ts: the STORED XLSX of
# xlsx_fixtures.ts encrypted by hand from [MS-OFFCRYPTO] (this script's
# own key derivation; AES by the openssl command), in a compound file; as
# base64. AGILE: Agile encryption, AES-256-CBC, SHA-512, a spin count of
# 1000 (Excel uses 100000; the file says which). STANDARD: Standard
# encryption, AES-128-ECB, SHA-1, its fixed 50000. Password: "Secret1".
import base64, hashlib, struct, subprocess, re, sys
sys.path.insert(0, __file__.rsplit('/', 1)[0])
from mkxls import cfb
src = open(__file__.rsplit('/', 1)[0] + '/xlsx_fixtures.ts').read()
plain = base64.b64decode(re.search(r'export const STORED = "([^"]+)"', src).group(1))
PW = "Secret1".encode('utf-16-le')
def aes(mode, key, data, iv=None):
    args = ['openssl', 'enc', f'-aes-{len(key)*8}-{mode}', '-K', key.hex(), '-nopad']
    if iv is not None: args += ['-iv', iv.hex()]
    return subprocess.run(args, input=data, capture_output=True, check=True).stdout
pad16 = lambda b: b + b'\x00' * (-len(b) % 16)
# --- Agile
salt = bytes(range(16)); ksalt = bytes(range(100, 116)); secret = bytes(range(200, 232)); vin = bytes(range(50, 66))
h = hashlib.sha512(salt + PW).digest()
for i in range(1000): h = hashlib.sha512(struct.pack('<I', i) + h).digest()
kf = lambda blk: hashlib.sha512(h + bytes(blk)).digest()[:32]
evhi = aes('cbc', kf([0xfe,0xa7,0xd2,0x76,0x3b,0x4b,0x9e,0x79]), vin, salt)
evhv = aes('cbc', kf([0xd7,0xaa,0x0f,0x6d,0x30,0x61,0x34,0x4e]), hashlib.sha512(vin).digest(), salt)
ekv = aes('cbc', kf([0x14,0x6e,0x0b,0xe7,0xab,0xac,0xd0,0xd6]), secret, salt)
enc = b''
for seg in range(0, len(plain), 4096):
    iv = hashlib.sha512(ksalt + struct.pack('<I', seg // 4096)).digest()[:16]
    enc += aes('cbc', secret, pad16(plain[seg:seg + 4096]), iv)
b64 = lambda b: base64.b64encode(b).decode()
xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{b64(ksalt)}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="1000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{b64(salt)}" encryptedVerifierHashInput="{b64(evhi)}" encryptedVerifierHashValue="{b64(evhv)}" encryptedKeyValue="{b64(ekv)}"/></keyEncryptor></keyEncryptors></encryption>'''.encode()
agile = cfb([('EncryptionInfo', struct.pack('<HHI', 4, 4, 0x40) + xml), ('EncryptedPackage', struct.pack('<Q', len(plain)) + enc)])
# --- Standard
ssalt = bytes(range(30, 46)); ver = bytes(range(60, 76))
h = hashlib.sha1(ssalt + PW).digest()
for i in range(50000): h = hashlib.sha1(struct.pack('<I', i) + h).digest()
h = hashlib.sha1(h + struct.pack('<I', 0)).digest()
x1 = hashlib.sha1(bytes(a ^ b for a, b in zip(h + b'\x00' * 44, b'\x36' * 64))).digest()
x2 = hashlib.sha1(bytes(a ^ b for a, b in zip(h + b'\x00' * 44, b'\x5c' * 64))).digest()
key = (x1 + x2)[:16]
header = struct.pack('<IIIIIIII', 0x24, 0, 0x660E, 0x8004, 128, 0x18, 0, 0) + 'Microsoft Enhanced RSA and AES Cryptographic Provider\x00'.encode('utf-16-le')
verifier = struct.pack('<I', 16) + ssalt + aes('ecb', key, ver) + struct.pack('<I', 20) + aes('ecb', key, pad16(hashlib.sha1(ver).digest()) + b'\x00' * 0)[:32]
if len(verifier) != 4 + 16 + 16 + 4 + 32: verifier = struct.pack('<I', 16) + ssalt + aes('ecb', key, ver) + struct.pack('<I', 20) + aes('ecb', key, hashlib.sha1(ver).digest() + b'\x00' * 12)
info = struct.pack('<HHII', 3, 2, 0x24, len(header)) + header + verifier
standard = cfb([('EncryptionInfo', info), ('EncryptedPackage', struct.pack('<Q', len(plain)) + aes('ecb', key, pad16(plain)))])
open(__file__.rsplit('/', 1)[0] + '/crypt_fixtures.ts', 'w').write(f'''// xlsx_fixtures.ts's STORED encrypted by hand from [MS-OFFCRYPTO]
// (mkcrypt.py beside this file: its own key derivation, AES by the openssl
// command), as base64; the password "Secret1". AGILE: Agile encryption,
// AES-256-CBC with SHA-512, a spin count of 1000. STANDARD: Standard
// encryption, AES-128-ECB with SHA-1.
export const AGILE = "{b64(agile)}";
export const STANDARD = "{b64(standard)}";
''')
print(len(agile), len(standard))
