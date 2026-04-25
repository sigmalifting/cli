import path from "node:path";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import {
  normalizeProcessBundle,
  normalizeProgramBundle,
} from "./bundleOps.js";
import { parseXlsxBufferToBundle } from "./xlsx/importWorkbook.js";
import { exportBundleToXlsxBuffer } from "./xlsx/exportWorkbook.js";

const normalizeFsPath = (source) =>
  typeof source === "string" && source.startsWith("file://")
    ? fileURLToPath(source)
    : source;

const detectBundleKind = (bundle) => {
  if (bundle?.process && bundle?.program && Array.isArray(bundle?.exercises)) {
    return "process-import";
  }

  if (bundle?.program && Array.isArray(bundle?.exercises) && !bundle?.process) {
    return "program-import";
  }

  throw new Error(
    "Unsupported bundle shape. Expected program-import or process-import."
  );
};

const normalizeBundle = (bundle) => {
  const bundleKind = detectBundleKind(bundle);
  return {
    bundleKind,
    normalized:
      bundleKind === "process-import"
        ? normalizeProcessBundle(bundle)
        : normalizeProgramBundle(bundle),
  };
};

export const importXlsxBuffer = async (inputBuffer) => {
  const parsed = await parseXlsxBufferToBundle(inputBuffer);
  if (!parsed.success || !parsed.data) {
    throw new Error(parsed.error || "Failed to parse XLSX workbook");
  }

  const { bundleKind, normalized } = normalizeBundle(parsed.data);
  return {
    bundle_kind: bundleKind,
    bundle: normalized.bundle,
    warnings: normalized.warnings || [],
  };
};

export const importXlsxFile = async (source) => {
  if (!source || source === "-") {
    throw new Error("XLSX import requires a local file path or file:// URL.");
  }

  const buffer = await readFile(normalizeFsPath(source));
  return importXlsxBuffer(buffer);
};

export const exportXlsxBuffer = async (bundle) => {
  const { bundleKind, normalized } = normalizeBundle(bundle);
  const exportResult = await exportBundleToXlsxBuffer(normalized.bundle);

  return {
    bundle_kind: bundleKind,
    filename: exportResult.filename,
    buffer: Buffer.from(exportResult.buffer),
  };
};

export const exportXlsxFile = async (bundle, outputTarget) => {
  if (!outputTarget) {
    throw new Error("Missing required --out flag for XLSX export.");
  }

  const { bundle_kind, filename, buffer } = await exportXlsxBuffer(bundle);
  const normalizedTarget = normalizeFsPath(outputTarget);
  await mkdir(path.dirname(normalizedTarget), { recursive: true });
  await writeFile(normalizedTarget, buffer);

  return {
    bundle_kind,
    filename,
    output_path: outputTarget,
  };
};
