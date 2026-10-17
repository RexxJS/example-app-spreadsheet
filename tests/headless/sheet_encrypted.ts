// Encrypted XLSX (5.7): SHA-1 and SHA-512 on FIPS 180-4's examples, AES on
// FIPS 197's; files encrypted by hand (lib/crypt_fixtures.ts) opened by
// their password, Agile (AES-256, SHA-512) and Standard (AES-128, SHA-1);
// a wrong password, and none, refused; OPENXLSX's password from REXX.
import { decryptBlock, decryptCbc, expandKey, hex, sha1, sha512 } from "./sheet/crypto.ts";
import { utf8Bytes } from "./sheet/zip.ts";
import { AGILE, STANDARD } from "./lib/crypt_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const bytes = (h: string) => new Uint8Array(h.match(/../g)!.map((x) => parseInt(x, 16)));
eq("SHA-1", [hex(sha1(utf8Bytes("abc"))), hex(sha1(utf8Bytes("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq")))],
  ["a9993e364706816aba3e25717850c26c9cd0d89d", "84983e441c3bd26ebaae4aa1f95129e5e54670f1"]);
eq("SHA-512", [hex(sha512(utf8Bytes("abc"))), hex(sha512(new Uint8Array(0))).slice(0, 32), hex(sha512(new Uint8Array(200).fill(0x61))).slice(0, 16)],
  ["ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f", "cf83e1357eefb8bdf1542850d66d8007", "4b11459c33f52a22"]);
eq("AES decryption: FIPS 197's 128- and 256-bit examples", [hex(decryptBlock(expandKey(bytes("000102030405060708090a0b0c0d0e0f")), bytes("69c4e0d86a7b0430d8cdb78070b4c55a"))),
  hex(decryptBlock(expandKey(bytes("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f")), bytes("8ea2b7ca516745bfeafc49904b496089")))],
  ["00112233445566778899aabbccddeeff", "00112233445566778899aabbccddeeff"]);
eq("CBC: each block after the one before", hex(decryptCbc(bytes("2b7e151628aed2a6abf7158809cf4f3c"), bytes("000102030405060708090a0b0c0d0e0f"), bytes("7649abac8119b246cee98e9b12e9197d5086cb9b507219ee95db113a917678b2"))),
  "6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e51");
const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
eq("encrypted: no password, the wrong one", [await message(`file.importXlsx name=Locked.xlsx base64=${AGILE}`), await message(`file.importXlsx base64=${AGILE} password=nope`)],
  ["Locked.xlsx is encrypted: open it with its password", "the password is wrong"]);
eq("Agile: opened by its password", [await message(`file.importXlsx name=Locked.xlsx base64=${AGILE} password=Secret1`), show(wb.sheet.value("A1")), wb.dateSystem], ["opened Locked: 1 cells, Excel formulas", "1", 1904]);
eq("Standard: the wrong password, then the right one", [await message(`file.importXlsx base64=${STANDARD} password=secret1`), await message(`file.importXlsx base64=${STANDARD} password=Secret1`)],
  ["the password is wrong", "opened Workbook: 1 cells, Excel formulas"]);
const said: string[] = [];
await cmds.script(`SAY OPENXLSX("${AGILE}", "Again.xlsx", "Secret1")`, (l) => said.push(l));
eq("OPENXLSX with a password from REXX", said, ["opened Again: 1 cells, Excel formulas"]);

// expect: ok SHA-1
// expect: ok SHA-512
// expect: ok AES decryption: FIPS 197's 128- and 256-bit examples
// expect: ok CBC: each block after the one before
// expect: ok encrypted: no password, the wrong one
// expect: ok Agile: opened by its password
// expect: ok Standard: the wrong password, then the right one
// expect: ok OPENXLSX with a password from REXX
