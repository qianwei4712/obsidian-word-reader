import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const pkg = JSON.parse(fs.readFileSync(path.join(rootDir, "package.json"), "utf8"));
const distDir = path.join(rootDir, "dist");
const releaseDir = path.join(rootDir, "release");
const outputPath = path.join(releaseDir, `${pkg.name}-v${pkg.version}.zip`);
const releaseFiles = ["main.js", "manifest.json", "styles.css"];

const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n += 1) {
  let c = n;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  crcTable[n] = c >>> 0;
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date) {
  const year = Math.max(date.getFullYear(), 1980);
  const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { dosDate, dosTime };
}

function uint16(value) {
  const buffer = Buffer.alloc(2);
  buffer.writeUInt16LE(value);
  return buffer;
}

function uint32(value) {
  const buffer = Buffer.alloc(4);
  buffer.writeUInt32LE(value >>> 0);
  return buffer;
}

function localFileHeader(entry) {
  return Buffer.concat([
    uint32(0x04034b50),
    uint16(20),
    uint16(0),
    uint16(0),
    uint16(entry.dosTime),
    uint16(entry.dosDate),
    uint32(entry.crc),
    uint32(entry.data.length),
    uint32(entry.data.length),
    uint16(entry.name.length),
    uint16(0),
    entry.name,
  ]);
}

function centralDirectoryHeader(entry) {
  return Buffer.concat([
    uint32(0x02014b50),
    uint16(20),
    uint16(20),
    uint16(0),
    uint16(0),
    uint16(entry.dosTime),
    uint16(entry.dosDate),
    uint32(entry.crc),
    uint32(entry.data.length),
    uint32(entry.data.length),
    uint16(entry.name.length),
    uint16(0),
    uint16(0),
    uint16(0),
    uint16(0),
    uint32(0),
    uint32(entry.offset),
    entry.name,
  ]);
}

function endOfCentralDirectory(entryCount, centralDirectorySize, centralDirectoryOffset) {
  return Buffer.concat([
    uint32(0x06054b50),
    uint16(0),
    uint16(0),
    uint16(entryCount),
    uint16(entryCount),
    uint32(centralDirectorySize),
    uint32(centralDirectoryOffset),
    uint16(0),
  ]);
}

for (const file of releaseFiles) {
  const fullPath = path.join(distDir, file);
  if (!fs.existsSync(fullPath)) {
    console.error(`Missing dist file: ${path.relative(rootDir, fullPath)}`);
    process.exit(1);
  }
}

fs.mkdirSync(releaseDir, { recursive: true });

const fileParts = [];
const entries = [];
let offset = 0;

for (const file of releaseFiles) {
  const fullPath = path.join(distDir, file);
  const data = fs.readFileSync(fullPath);
  const stats = fs.statSync(fullPath);
  const { dosDate, dosTime } = dosDateTime(stats.mtime);
  const entry = {
    crc: crc32(data),
    data,
    dosDate,
    dosTime,
    name: Buffer.from(file, "utf8"),
    offset,
  };
  const header = localFileHeader(entry);
  fileParts.push(header, data);
  entries.push(entry);
  offset += header.length + data.length;
}

const centralDirectoryOffset = offset;
const centralDirectoryParts = entries.map(centralDirectoryHeader);
const centralDirectorySize = centralDirectoryParts.reduce((total, part) => total + part.length, 0);
const archive = Buffer.concat([
  ...fileParts,
  ...centralDirectoryParts,
  endOfCentralDirectory(entries.length, centralDirectorySize, centralDirectoryOffset),
]);

fs.writeFileSync(outputPath, archive);
console.log(`Created ${path.relative(rootDir, outputPath)} with ${releaseFiles.join(", ")}.`);
