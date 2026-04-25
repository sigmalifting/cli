import os from "node:os";
import path from "node:path";
import {
  mkdir,
  readFile,
  readdir,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import {
  normalizeProcessBundle,
  normalizeProgramBundle,
} from "./bundleOps.js";

const INDEX_VERSION = 1;

const normalizeRoot = (root) => path.resolve(root);

const getDefaultStoreRoot = () => {
  if (process.env.SIGMALIFTING_HOME) {
    return normalizeRoot(process.env.SIGMALIFTING_HOME);
  }

  if (process.platform === "darwin") {
    return path.join(
      os.homedir(),
      "Library",
      "Application Support",
      "SigmaLifting",
      "cli"
    );
  }

  if (process.platform === "win32") {
    return path.join(
      process.env.APPDATA || path.join(os.homedir(), "AppData", "Roaming"),
      "SigmaLifting",
      "cli"
    );
  }

  return path.join(
    process.env.XDG_DATA_HOME || path.join(os.homedir(), ".local", "share"),
    "sigmalifting",
    "cli"
  );
};

export const getStorePaths = ({ root } = {}) => {
  const storeRoot = normalizeRoot(root || getDefaultStoreRoot());

  return {
    root: storeRoot,
    programs_dir: path.join(storeRoot, "programs"),
    processes_dir: path.join(storeRoot, "processes"),
    exports_dir: path.join(storeRoot, "exports"),
    index_path: path.join(storeRoot, "index.json"),
  };
};

const ensureStoreDirectories = async (paths) => {
  await mkdir(paths.programs_dir, { recursive: true });
  await mkdir(paths.processes_dir, { recursive: true });
  await mkdir(paths.exports_dir, { recursive: true });
};

const slugifyName = (name, fallback) => {
  const slug = String(name || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-")
    .slice(0, 60);

  return slug || fallback;
};

export const detectBundleKind = (bundle) => {
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

const normalizeBundleForStore = (bundle) => {
  const bundleKind = detectBundleKind(bundle);
  const normalized =
    bundleKind === "process-import"
      ? normalizeProcessBundle(bundle)
      : normalizeProgramBundle(bundle);

  return {
    bundle_kind: bundleKind,
    entity_type: bundleKind === "process-import" ? "process" : "program",
    bundle: normalized.bundle,
    warnings: normalized.warnings || [],
  };
};

const getEntityIdentity = (bundle, entityType) => {
  if (entityType === "process") {
    return {
      id: bundle.process._id,
      name: bundle.process.name,
      fallbackSlug: "process",
    };
  }

  return {
    id: bundle.program._id,
    name: bundle.program.name,
    fallbackSlug: "program",
  };
};

const getEntityDir = (paths, entityType) =>
  entityType === "process" ? paths.processes_dir : paths.programs_dir;

const buildEntityPath = (paths, bundle, entityType) => {
  const { id, name, fallbackSlug } = getEntityIdentity(bundle, entityType);
  const slug = slugifyName(name, fallbackSlug);
  return path.join(getEntityDir(paths, entityType), `${slug}__${id}.json`);
};

const serializeJson = (value) => `${JSON.stringify(value, null, 2)}\n`;

const writeJsonAtomic = async (targetPath, value) => {
  await mkdir(path.dirname(targetPath), { recursive: true });
  const tempPath = `${targetPath}.${process.pid}.${Date.now()}.tmp`;

  try {
    await writeFile(tempPath, serializeJson(value), "utf8");
    await rename(tempPath, targetPath);
  } catch (error) {
    await rm(tempPath, { force: true });
    throw error;
  }
};

const readJsonFile = async (filePath) =>
  JSON.parse(await readFile(filePath, "utf8"));

const listJsonFiles = async (dir) => {
  try {
    const names = await readdir(dir);
    return names
      .filter((name) => name.endsWith(".json") && !name.includes(".tmp"))
      .map((name) => path.join(dir, name));
  } catch (error) {
    if (error.code === "ENOENT") {
      return [];
    }
    throw error;
  }
};

const countCompletedSets = (process) => {
  let completed_sets = 0;
  let total_sets = 0;

  (process.exercise_recordings || []).forEach((recording) => {
    (recording.weekly || []).forEach((week) => {
      (week.sets || []).forEach((set) => {
        total_sets += 1;
        if (set.completed) {
          completed_sets += 1;
        }
      });
    });
  });

  return {
    completed_sets,
    total_sets,
  };
};

const buildSummary = (bundle, entityType, filePath, root) => {
  if (entityType === "process") {
    const setCounts = countCompletedSets(bundle.process);

    return {
      id: bundle.process._id,
      kind: "process-import",
      entity_type: "process",
      name: bundle.process.name,
      program_id: bundle.process.program_id,
      program_name: bundle.process.program_name,
      start_date: bundle.process.start_date,
      created_at: bundle.process.created_at,
      updated_at: bundle.process.updated_at,
      exercise_recordings: bundle.process.exercise_recordings?.length || 0,
      ...setCounts,
      path: filePath,
      relative_path: path.relative(root, filePath),
    };
  }

  return {
    id: bundle.program._id,
    kind: "program-import",
    entity_type: "program",
    name: bundle.program.name,
    created_at: bundle.program.created_at,
    updated_at: bundle.program.updated_at,
    blocks: bundle.program.blocks?.length || 0,
    exercises: bundle.exercises?.length || 0,
    path: filePath,
    relative_path: path.relative(root, filePath),
  };
};

const scanEntityDir = async (paths, entityType) => {
  const dir = getEntityDir(paths, entityType);
  const files = await listJsonFiles(dir);
  const entries = [];
  const errors = [];

  for (const filePath of files) {
    try {
      const source = await readJsonFile(filePath);
      const normalized = normalizeBundleForStore(source);

      if (normalized.entity_type !== entityType) {
        throw new Error(
          `Expected ${entityType} bundle but found ${normalized.entity_type}`
        );
      }

      entries.push(buildSummary(normalized.bundle, entityType, filePath, paths.root));
    } catch (error) {
      errors.push({
        path: filePath,
        message: error.message || String(error),
      });
    }
  }

  entries.sort((left, right) => {
    const updated = String(right.updated_at || "").localeCompare(
      String(left.updated_at || "")
    );
    return updated || left.name.localeCompare(right.name);
  });

  return { entries, errors };
};

export const rebuildStoreIndex = async ({ root } = {}) => {
  const paths = getStorePaths({ root });
  await ensureStoreDirectories(paths);

  const programs = await scanEntityDir(paths, "program");
  const processes = await scanEntityDir(paths, "process");
  const index = {
    version: INDEX_VERSION,
    updated_at: new Date().toISOString(),
    root: paths.root,
    programs: programs.entries,
    processes: processes.entries,
    errors: [...programs.errors, ...processes.errors],
  };

  await writeJsonAtomic(paths.index_path, index);
  return {
    store: paths,
    index,
  };
};

export const initStore = async ({ root } = {}) => rebuildStoreIndex({ root });

const removeStaleEntityFiles = async (paths, entityType, id, canonicalPath) => {
  const files = await listJsonFiles(getEntityDir(paths, entityType));
  const suffix = `__${id}.json`;

  await Promise.all(
    files
      .filter((filePath) => filePath.endsWith(suffix) && filePath !== canonicalPath)
      .map((filePath) => rm(filePath, { force: true }))
  );
};

export const saveBundleToStore = async (bundle, { root } = {}) => {
  const paths = getStorePaths({ root });
  await ensureStoreDirectories(paths);

  const normalized = normalizeBundleForStore(bundle);
  const targetPath = buildEntityPath(
    paths,
    normalized.bundle,
    normalized.entity_type
  );
  const identity = getEntityIdentity(normalized.bundle, normalized.entity_type);

  await removeStaleEntityFiles(
    paths,
    normalized.entity_type,
    identity.id,
    targetPath
  );
  await writeJsonAtomic(targetPath, normalized.bundle);

  const { index } = await rebuildStoreIndex({ root: paths.root });

  return {
    root: paths.root,
    path: targetPath,
    relative_path: path.relative(paths.root, targetPath),
    index_path: paths.index_path,
    kind: normalized.bundle_kind,
    entity_type: normalized.entity_type,
    id: identity.id,
    name: identity.name,
    warnings: normalized.warnings,
    index_errors: index.errors,
  };
};

export const listStoredBundles = async (entityType, { root } = {}) => {
  const paths = getStorePaths({ root });
  const { index } = await rebuildStoreIndex({ root: paths.root });

  return {
    store: paths,
    entries: entityType === "process" ? index.processes : index.programs,
    errors: index.errors,
  };
};

const findStoredBundlePath = async (paths, entityType, id) => {
  if (!id || typeof id !== "string") {
    throw new Error(`Missing required ${entityType} id`);
  }

  const files = await listJsonFiles(getEntityDir(paths, entityType));
  const suffix = `__${id}.json`;
  const filenameMatch = files.find((filePath) => filePath.endsWith(suffix));

  if (filenameMatch) {
    return filenameMatch;
  }

  for (const filePath of files) {
    try {
      const source = await readJsonFile(filePath);
      const normalized = normalizeBundleForStore(source);
      const identity = getEntityIdentity(normalized.bundle, normalized.entity_type);

      if (normalized.entity_type === entityType && identity.id === id) {
        return filePath;
      }
    } catch (_error) {
      // Invalid files are surfaced by list/index commands; lookup keeps searching.
    }
  }

  throw new Error(`No stored ${entityType} found for id: ${id}`);
};

export const loadStoredBundle = async (entityType, id, { root } = {}) => {
  const paths = getStorePaths({ root });
  await ensureStoreDirectories(paths);

  const filePath = await findStoredBundlePath(paths, entityType, id);
  const source = await readJsonFile(filePath);
  const normalized = normalizeBundleForStore(source);

  if (normalized.entity_type !== entityType) {
    throw new Error(
      `Stored file for ${id} is ${normalized.entity_type}, not ${entityType}`
    );
  }

  return {
    bundle: normalized.bundle,
    storage: {
      root: paths.root,
      path: filePath,
      relative_path: path.relative(paths.root, filePath),
      kind: normalized.bundle_kind,
      entity_type: normalized.entity_type,
      id,
      name: getEntityIdentity(normalized.bundle, entityType).name,
      index_path: paths.index_path,
    },
  };
};
