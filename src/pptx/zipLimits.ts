export interface ZipSafetyLimits {
  maxFileCount: number;
  maxEntryUncompressedBytes: number;
  maxTotalUncompressedBytes: number;
  maxCompressionRatio: number;
}

export const DEFAULT_PPTX_ZIP_LIMITS: ZipSafetyLimits = {
  maxFileCount: 2_000,
  maxEntryUncompressedBytes: 64 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxCompressionRatio: 200,
};

export type ZipSafetyErrorKind =
  | "format-mismatch"
  | "encrypted"
  | "damaged"
  | "limit-exceeded";

export class ZipSafetyError extends Error {
  constructor(
    readonly kind: ZipSafetyErrorKind,
    message: string,
  ) {
    super(message);
    this.name = "ZipSafetyError";
  }
}

export interface ZipSafetySummary {
  fileCount: number;
  totalCompressedBytes: number;
  totalUncompressedBytes: number;
}

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE = 0x07064b50;
const MAX_END_RECORD_SEARCH = 65_535 + 22;
const OLE_COMPOUND_FILE_SIGNATURE = [
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
];

export function validateZipSafety(
  buffer: ArrayBuffer,
  limits: ZipSafetyLimits = DEFAULT_PPTX_ZIP_LIMITS,
): ZipSafetySummary {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  if (hasPrefix(bytes, OLE_COMPOUND_FILE_SIGNATURE)) {
    throw new ZipSafetyError(
      "encrypted",
      "Encrypted or legacy binary PowerPoint containers cannot be previewed.",
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
      "ZIP64 PowerPoint packages are not supported by the safe preview limits.",
    );
  }

  const diskNumber = view.getUint16(endOffset + 4, true);
  const centralDirectoryDisk = view.getUint16(endOffset + 6, true);
  const fileCount = view.getUint16(endOffset + 10, true);
  const centralDirectorySize = view.getUint32(endOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(endOffset + 16, true);
  if (diskNumber !== 0 || centralDirectoryDisk !== 0) {
    throw new ZipSafetyError(
      "limit-exceeded",
      "Multi-disk PowerPoint packages are not supported.",
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
    if ((generalPurposeFlags & 0x1) !== 0) {
      throw new ZipSafetyError(
        "encrypted",
        "The PowerPoint package contains encrypted ZIP entries.",
      );
    }

    const compressedBytes = view.getUint32(offset + 20, true);
    const uncompressedBytes = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraFieldLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    if (
      compressedBytes === 0xffffffff ||
      uncompressedBytes === 0xffffffff
    ) {
      throw new ZipSafetyError(
        "limit-exceeded",
        "ZIP64 entries are not supported by the safe preview limits.",
      );
    }
    if (uncompressedBytes > limits.maxEntryUncompressedBytes) {
      throw new ZipSafetyError(
        "limit-exceeded",
        `A package entry expands to ${uncompressedBytes} bytes, exceeding the per-entry limit.`,
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
        `A package entry has an unsafe compression ratio of ${compressionRatio.toFixed(1)}.`,
      );
    }

    totalCompressedBytes += compressedBytes;
    totalUncompressedBytes += uncompressedBytes;
    if (totalUncompressedBytes > limits.maxTotalUncompressedBytes) {
      throw new ZipSafetyError(
        "limit-exceeded",
        `The package expands to more than ${limits.maxTotalUncompressedBytes} bytes.`,
      );
    }

    offset += 46 + fileNameLength + extraFieldLength + commentLength;
    if (offset > centralDirectoryOffset + centralDirectorySize) {
      throw new ZipSafetyError(
        "damaged",
        "The ZIP central directory entry lengths are inconsistent.",
      );
    }
  }

  return {
    fileCount,
    totalCompressedBytes,
    totalUncompressedBytes,
  };
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
