export interface ZipSafetyLimits {
  maxFileCount: number;
  maxEntryUncompressedBytes: number;
  maxTotalUncompressedBytes: number;
  maxCompressionRatio: number;
}

export const OOXML_ZIP_HARD_LIMITS = Object.freeze<ZipSafetyLimits>({
  maxFileCount: 10_000,
  maxEntryUncompressedBytes: 128 * 1024 * 1024,
  maxTotalUncompressedBytes: 512 * 1024 * 1024,
  maxCompressionRatio: 200,
});

export const DEFAULT_PPTX_ZIP_LIMITS = Object.freeze<ZipSafetyLimits>({
  maxFileCount: 2_000,
  maxEntryUncompressedBytes: 64 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxCompressionRatio: 200,
});

export const DEFAULT_DOCX_ZIP_LIMITS = Object.freeze<ZipSafetyLimits>({
  maxFileCount: 2_000,
  maxEntryUncompressedBytes: 64 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxCompressionRatio: 200,
});

export const DEFAULT_XLSX_ZIP_LIMITS = Object.freeze<ZipSafetyLimits>({
  ...OOXML_ZIP_HARD_LIMITS,
});

export type ZipSafetyErrorKind =
  | "format-mismatch"
  | "encrypted"
  | "damaged"
  | "limit-exceeded"
  | "unsupported";

export class ZipSafetyError extends Error {
  constructor(
    readonly kind: ZipSafetyErrorKind,
    message: string,
  ) {
    super(message);
    this.name = "ZipSafetyError";
  }
}

export interface ZipSafetyEntry {
  name: string;
  compressedBytes: number;
  uncompressedBytes: number;
  compressionMethod: number;
}

export interface ZipSafetySummary {
  fileCount: number;
  totalCompressedBytes: number;
  totalUncompressedBytes: number;
  entries: readonly ZipSafetyEntry[];
}

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE = 0x07064b50;
const MAX_END_RECORD_SEARCH = 65_535 + 22;
const OLE_COMPOUND_FILE_SIGNATURE = [
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
];
const SUPPORTED_COMPRESSION_METHODS = new Set([0, 8]);

export function validateZipSafety(
  buffer: ArrayBuffer,
  limits: ZipSafetyLimits = OOXML_ZIP_HARD_LIMITS,
): ZipSafetySummary {
  assertLimitsWithinHardCaps(limits);

  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  if (hasPrefix(bytes, OLE_COMPOUND_FILE_SIGNATURE)) {
    throw new ZipSafetyError(
      "encrypted",
      "Encrypted or legacy binary Office containers cannot be previewed.",
    );
  }
  if (bytes.length < 22) {
    throw new ZipSafetyError(
      "format-mismatch",
      "The file is too small to be a ZIP-based Office document.",
    );
  }

  const endOffset = findEndOfCentralDirectory(view);
  if (endOffset < 0) {
    throw new ZipSafetyError(
      "format-mismatch",
      "The ZIP end-of-central-directory record is missing.",
    );
  }

  if (
    endOffset >= 20 &&
    view.getUint32(endOffset - 20, true) ===
      ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE
  ) {
    throw new ZipSafetyError(
      "limit-exceeded",
      "ZIP64 Office packages exceed the supported safe preview envelope.",
    );
  }

  const diskNumber = view.getUint16(endOffset + 4, true);
  const centralDirectoryDisk = view.getUint16(endOffset + 6, true);
  const diskFileCount = view.getUint16(endOffset + 8, true);
  const fileCount = view.getUint16(endOffset + 10, true);
  const centralDirectorySize = view.getUint32(endOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(endOffset + 16, true);
  if (diskNumber !== 0 || centralDirectoryDisk !== 0) {
    throw new ZipSafetyError(
      "limit-exceeded",
      "Multi-disk Office packages are not supported.",
    );
  }
  if (
    diskFileCount === 0xffff ||
    fileCount === 0xffff ||
    centralDirectorySize === 0xffffffff ||
    centralDirectoryOffset === 0xffffffff
  ) {
    throw new ZipSafetyError(
      "limit-exceeded",
      "ZIP64 Office packages exceed the supported safe preview envelope.",
    );
  }
  if (diskFileCount !== fileCount) {
    throw new ZipSafetyError(
      "damaged",
      "The ZIP central-directory entry counts are inconsistent.",
    );
  }
  if (fileCount > limits.maxFileCount) {
    throw new ZipSafetyError(
      "limit-exceeded",
      `The package contains ${fileCount} files, exceeding the ${limits.maxFileCount} file limit.`,
    );
  }
  if (
    centralDirectoryOffset + centralDirectorySize > endOffset ||
    centralDirectoryOffset > bytes.length
  ) {
    throw new ZipSafetyError(
      "damaged",
      "The ZIP central directory points outside the file.",
    );
  }

  const entries: ZipSafetyEntry[] = [];
  const entryNames = new Set<string>();
  let offset = centralDirectoryOffset;
  let totalCompressedBytes = 0;
  let totalUncompressedBytes = 0;
  for (let index = 0; index < fileCount; index += 1) {
    if (
      offset + 46 > bytes.length ||
      view.getUint32(offset, true) !== CENTRAL_DIRECTORY_SIGNATURE
    ) {
      throw new ZipSafetyError(
        "damaged",
        "The ZIP central directory is incomplete or malformed.",
      );
    }

    const generalPurposeFlags = view.getUint16(offset + 8, true);
    if ((generalPurposeFlags & 0x41) !== 0) {
      throw new ZipSafetyError(
        "encrypted",
        "The Office package contains encrypted ZIP entries.",
      );
    }

    const compressionMethod = view.getUint16(offset + 10, true);
    if (!SUPPORTED_COMPRESSION_METHODS.has(compressionMethod)) {
      throw new ZipSafetyError(
        "unsupported",
        `ZIP compression method ${compressionMethod} is not supported.`,
      );
    }

    const compressedBytes = view.getUint32(offset + 20, true);
    const uncompressedBytes = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraFieldLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    if (
      compressedBytes === 0xffffffff ||
      uncompressedBytes === 0xffffffff ||
      localHeaderOffset === 0xffffffff
    ) {
      throw new ZipSafetyError(
        "limit-exceeded",
        "ZIP64 entries exceed the supported safe preview envelope.",
      );
    }

    const entryEnd =
      offset + 46 + fileNameLength + extraFieldLength + commentLength;
    if (
      entryEnd > bytes.length ||
      entryEnd > centralDirectoryOffset + centralDirectorySize
    ) {
      throw new ZipSafetyError(
        "damaged",
        "The ZIP central-directory entry lengths are inconsistent.",
      );
    }

    const nameBytes = bytes.subarray(offset + 46, offset + 46 + fileNameLength);
    const name = new TextDecoder("utf-8").decode(nameBytes);
    validateEntryName(name);
    if (entryNames.has(name)) {
      throw new ZipSafetyError(
        "damaged",
        `The ZIP package contains a duplicate entry named ${name}.`,
      );
    }
    entryNames.add(name);
    validateLocalHeader(
      view,
      bytes,
      localHeaderOffset,
      compressedBytes,
      centralDirectoryOffset,
    );

    if (uncompressedBytes > limits.maxEntryUncompressedBytes) {
      throw new ZipSafetyError(
        "limit-exceeded",
        `The package entry ${name} expands past the per-entry limit.`,
      );
    }
    const compressionRatio =
      uncompressedBytes / Math.max(compressedBytes, 1);
    if (
      uncompressedBytes > 0 &&
      compressionRatio > limits.maxCompressionRatio
    ) {
      throw new ZipSafetyError(
        "limit-exceeded",
        `The package entry ${name} has an unsafe compression ratio of ${compressionRatio.toFixed(1)}.`,
      );
    }

    totalCompressedBytes += compressedBytes;
    totalUncompressedBytes += uncompressedBytes;
    if (totalUncompressedBytes > limits.maxTotalUncompressedBytes) {
      throw new ZipSafetyError(
        "limit-exceeded",
        `The package expands past the ${limits.maxTotalUncompressedBytes} byte total limit.`,
      );
    }

    entries.push({
      name,
      compressedBytes,
      uncompressedBytes,
      compressionMethod,
    });
    offset = entryEnd;
  }

  if (offset !== centralDirectoryOffset + centralDirectorySize) {
    throw new ZipSafetyError(
      "damaged",
      "The ZIP central-directory size is inconsistent with its entries.",
    );
  }

  return {
    fileCount,
    totalCompressedBytes,
    totalUncompressedBytes,
    entries,
  };
}

function assertLimitsWithinHardCaps(limits: ZipSafetyLimits): void {
  for (const [key, hardCap] of Object.entries(OOXML_ZIP_HARD_LIMITS)) {
    const value = limits[key as keyof ZipSafetyLimits];
    if (!Number.isFinite(value) || value <= 0 || value > hardCap) {
      throw new RangeError(
        `${key} must be positive and no greater than the global OOXML hard cap ${hardCap}.`,
      );
    }
  }
}

function validateEntryName(name: string): void {
  const parts = name.split("/");
  if (
    name.length === 0 ||
    name.includes("\\") ||
    name.includes("\0") ||
    name.startsWith("/") ||
    /^[a-z]:/i.test(name) ||
    parts.some((part) => part === "..")
  ) {
    throw new ZipSafetyError(
      "damaged",
      "The ZIP package contains an unsafe entry path.",
    );
  }
}

function validateLocalHeader(
  view: DataView,
  bytes: Uint8Array,
  offset: number,
  compressedBytes: number,
  centralDirectoryOffset: number,
): void {
  if (
    offset + 30 > centralDirectoryOffset ||
    view.getUint32(offset, true) !== LOCAL_FILE_HEADER_SIGNATURE
  ) {
    throw new ZipSafetyError(
      "damaged",
      "A ZIP central-directory entry points to an invalid local header.",
    );
  }
  const flags = view.getUint16(offset + 6, true);
  if ((flags & 0x41) !== 0) {
    throw new ZipSafetyError(
      "encrypted",
      "The Office package contains encrypted ZIP entries.",
    );
  }
  const nameLength = view.getUint16(offset + 26, true);
  const extraLength = view.getUint16(offset + 28, true);
  const dataEnd = offset + 30 + nameLength + extraLength + compressedBytes;
  if (dataEnd > centralDirectoryOffset || dataEnd > bytes.length) {
    throw new ZipSafetyError(
      "damaged",
      "A ZIP entry payload points outside the package data region.",
    );
  }
}

function hasPrefix(bytes: Uint8Array, prefix: number[]): boolean {
  return (
    bytes.length >= prefix.length &&
    prefix.every((value, index) => bytes[index] === value)
  );
}

function findEndOfCentralDirectory(view: DataView): number {
  const minimumOffset = Math.max(0, view.byteLength - MAX_END_RECORD_SEARCH);
  for (let offset = view.byteLength - 22; offset >= minimumOffset; offset -= 1) {
    if (
      view.getUint32(offset, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE
    ) {
      const commentLength = view.getUint16(offset + 20, true);
      if (offset + 22 + commentLength === view.byteLength) {
        return offset;
      }
    }
  }
  return -1;
}
