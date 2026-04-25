import path from "node:path";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import {
  SCHEMA_CATALOG,
  addBlockCommand,
  addCustomLiftCommand,
  addDayCommand,
  addExerciseCommand,
  copyBlockCommand,
  copyDayCommand,
  copyExerciseCommand,
  copyProgramCommand,
  createProcessFromProgramCommand,
  createProgramCommand,
  createTemplate,
  deleteBlockCommand,
  deleteCustomLiftCommand,
  deleteDayCommand,
  deleteExerciseCommand,
  getExerciseRecordingCommand,
  moveExerciseCommand,
  normalizeProcessBundle,
  normalizeProgramBundle,
  renameCustomLiftCommand,
  renameDayCommand,
  updateBlockCommand,
  updateExerciseCommand,
  updateProcessCommand,
  updateProcessConfigAndOneRmCommand,
  updateProcessConfigCommand,
  updateProcessNoteCommand,
  updateProcessOneRmCommand,
  updateProcessSetCommand,
  updateProgramCommand,
  updateScheduleCommand,
  validatePayload,
} from "./lib/bundleOps.js";
import {
  INSPECT_VIEWS,
  RENDER_VIEWS,
  buildInspectPage,
  renderInspectPage,
} from "./lib/inspectOps.js";
import {
  getStorePaths,
  initStore,
  listStoredBundles,
  loadStoredBundle,
  saveBundleToStore,
} from "./lib/storeOps.js";
import {
  RPE_CHART_DATA,
  calculate1RM,
  calculateSuggestedWeightRange,
  calculateTrainingWeight,
} from "./lib/rpe.js";
import { exportXlsxFile, importXlsxFile } from "./lib/xlsxOps.js";

const ROOT_COMMANDS = [
  "help",
  "doctor",
  "template",
  "validate",
  "normalize",
  "program",
  "process",
  "rpe",
  "schema",
  "inspect",
  "render",
  "xlsx",
  "store",
];
const PROGRAM_SUBCOMMANDS = [
  "list",
  "show",
  "create",
  "copy",
  "update",
  "add-block",
  "update-block",
  "delete-block",
  "add-day",
  "rename-day",
  "delete-day",
  "update-schedule",
  "add-exercise",
  "update-exercise",
  "delete-exercise",
  "move-exercise",
  "copy-exercise",
  "copy-block",
  "copy-day",
  "add-custom-lift",
  "rename-custom-lift",
  "delete-custom-lift",
];
const PROCESS_SUBCOMMANDS = [
  "list",
  "show",
  "create-from-program",
  "update",
  "update-config",
  "update-one-rm",
  "update-config-and-one-rm",
  "update-set",
  "update-note",
  "get-recording",
];
const RPE_SUBCOMMANDS = ["chart", "e1rm", "training-weight", "range"];
const SCHEMA_SUBCOMMANDS = ["list", "show"];
const XLSX_SUBCOMMANDS = ["import", "export"];
const STORE_SUBCOMMANDS = ["path", "init"];

const HELP_TEXT = `SigmaLifting agent CLI

Usage:
  sigmalifting-cli help
  sigmalifting-cli doctor
  sigmalifting-cli template <kind>
  sigmalifting-cli validate <kind> --file <path|"-">
  sigmalifting-cli normalize <program-import|process-import> --file <path|"-">
  sigmalifting-cli program <command> [options]
  sigmalifting-cli process <command> [options]
  sigmalifting-cli rpe <chart|e1rm|training-weight|range> [options]
  sigmalifting-cli schema <list|show> [kind]
  sigmalifting-cli inspect <view> --file <path|"-">
  sigmalifting-cli render <view> --file <path|"-"> [--raw]
  sigmalifting-cli xlsx <import|export> [options]
  sigmalifting-cli store <path|init>

Program commands:
  list
  show --program-id <id>
  create
  copy --file <program.json>
  update --file <program.json>|--program-id <id> --data '{...}'
  add-block --file <program.json>|--program-id <id> --data '{...}'
  update-block --file <program.json>|--program-id <id> --block-id <id> --data '{...}'
  delete-block --file <program.json>|--program-id <id> --block-id <id>
  add-day --file <program.json>|--program-id <id> --block-id <id> --position <0-6> --data '{...}'
  rename-day --file <program.json>|--program-id <id> --block-id <id> --day-id <id> --name "New Name"
  delete-day --file <program.json>|--program-id <id> --block-id <id> --day-id <id>
  update-schedule --file <program.json>|--program-id <id> --block-id <id> --data '["dayA","","",...]'
  add-exercise --file <program.json>|--program-id <id> --block-id <id> --day-id <id> --data '{...}'
  update-exercise --file <program.json>|--program-id <id> --exercise-id <id> --data '{...}'
  delete-exercise --file <program.json>|--program-id <id> --block-id <id> --day-id <id> --exercise-id <id>
  move-exercise --file <program.json>|--program-id <id> --block-id <id> --day-id <id> --exercise-id <id> --direction <up|down|left|right>
  copy-exercise --file <program.json>|--program-id <id> --block-id <id> --day-id <id> --exercise-id <id>
  copy-block --file <program.json>|--program-id <id> --block-id <id>
  copy-day --file <program.json>|--program-id <id> --block-id <id> --day-id <id>
  add-custom-lift --file <program.json>|--program-id <id> --name "Front Squat"
  rename-custom-lift --file <program.json>|--program-id <id> --old-name "Front Squat" --new-name "Pause Squat"
  delete-custom-lift --file <program.json>|--program-id <id> --name "Front Squat"

Process commands:
  list
  show --process-id <id>
  create-from-program --file <program.json>|--program-id <id> [--name "Meet Prep"] [--start-date YYYY-MM-DD]
  update --file <process.json>|--process-id <id> --data '{...}'
  update-config --file <process.json>|--process-id <id> --data '{...}'
  update-one-rm --file <process.json>|--process-id <id> --data '{...}'
  update-config-and-one-rm --file <process.json>|--process-id <id> --config '{...}' --one-rm '{...}'
  update-set --file <process.json>|--process-id <id> --exercise-id <id> --week-index <n> --set-index <n> --data '{...}'
  update-note --file <process.json>|--process-id <id> --exercise-id <id> --week-index <n> --note "..."
  get-recording --file <process.json>|--process-id <id> --exercise-id <id>

XLSX commands:
  import --file <workbook.xlsx> [--out <bundle.json>]
  export --file <path|"-"> --out <workbook.xlsx>

Store commands:
  path
  init

Flags:
  --file <path|"-">  Read JSON input from a file or stdin
  --data '{...}'     Inline JSON payload
  --data-file <path> Read JSON payload from a file
  --out <path>       Write the resulting bundle or workbook to a file
  --exercise-id <id> Select an exercise for exercise-detail, recording-detail, and process recording commands
  --compact          Print compact JSON instead of pretty JSON
  --raw              For render, print raw text directly
  --store-root <dir> Override the local JSON store root for one invocation

Notes:
  --file, --data-file, and --out accept local paths and file:// URLs in Node
  Program and process bundle results are stored as local JSON by default
  Set SIGMALIFTING_HOME to override the local storage root
`;

const parseArgs = (argv) => {
  const positionals = [];
  const flags = {};

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (!arg.startsWith("--")) {
      positionals.push(arg);
      continue;
    }

    const key = arg.slice(2);
    const nextArg = argv[index + 1];

    if (!nextArg || nextArg.startsWith("--")) {
      flags[key] = true;
      continue;
    }

    flags[key] = nextArg;
    index += 1;
  }

  return { positionals, flags };
};

const serialize = (value, compact = false) =>
  `${JSON.stringify(value, null, compact ? 0 : 2)}\n`;

const defaultStdout = (value) => process.stdout.write(value);
const defaultStderr = (value) => process.stderr.write(value);

const normalizeFsPath = (source) =>
  typeof source === "string" && source.startsWith("file://")
    ? fileURLToPath(source)
    : source;

const readStdin = async () => {
  const chunks = [];
  for await (const chunk of process.stdin) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks).toString("utf8");
};

const readJsonSource = async (source, fallbackStdin) => {
  if (source === "-") {
    const text = fallbackStdin !== undefined ? fallbackStdin : await readStdin();
    return JSON.parse(text);
  }

  const text = await readFile(normalizeFsPath(source), "utf8");
  return JSON.parse(text);
};

const readBundle = async (flags, stdinText) => {
  const source = flags.file;
  if (!source) {
    throw new Error("Missing required --file flag");
  }
  return readJsonSource(source, stdinText);
};

const readPayload = async (flags, key = "data", stdinText) => {
  const inlineValue = flags[key];
  const fileValue = flags[`${key}-file`];

  if (inlineValue === "-") {
    const text = stdinText !== undefined ? stdinText : await readStdin();
    return JSON.parse(text);
  }

  if (inlineValue && inlineValue !== true) {
    return JSON.parse(inlineValue);
  }

  if (fileValue) {
    return readJsonSource(fileValue, stdinText);
  }

  throw new Error(`Missing required --${key} or --${key}-file flag`);
};

const writeBundleIfRequested = async (bundle, flags) => {
  if (!flags.out) {
    return null;
  }

  const outputTarget = normalizeFsPath(flags.out);
  await mkdir(path.dirname(outputTarget), { recursive: true });
  await writeFile(outputTarget, serialize(bundle, false), "utf8");
  return flags.out;
};

const getProgramIdFlag = (flags) =>
  typeof flags["program-id"] === "string"
    ? flags["program-id"]
    : typeof flags.id === "string"
      ? flags.id
      : null;

const getProcessIdFlag = (flags) =>
  typeof flags["process-id"] === "string"
    ? flags["process-id"]
    : typeof flags.id === "string"
      ? flags.id
      : null;

const readProgramBundle = async (flags, stdinText, storeOptions) => {
  const programId = getProgramIdFlag(flags);
  if (programId) {
    const stored = await loadStoredBundle("program", programId, storeOptions);
    return stored.bundle;
  }

  return readBundle(flags, stdinText);
};

const readProcessBundle = async (flags, stdinText, storeOptions) => {
  const processId = getProcessIdFlag(flags);
  if (processId) {
    const stored = await loadStoredBundle("process", processId, storeOptions);
    return stored.bundle;
  }

  return readBundle(flags, stdinText);
};

const buildSuccess = (command, payload, outputPath = null) => ({
  ok: true,
  command,
  ...payload,
  ...(outputPath ? { output_path: outputPath } : {}),
});

const buildError = (command, error) => ({
  ok: false,
  command,
  error: {
    message: error.message || String(error),
    details: error.details || null,
  },
});

const buildHelpPayload = () => ({
  summary: "Create, validate, normalize, and mutate SigmaLifting JSON bundles.",
  usage: [
    "sigmalifting-cli help",
    "sigmalifting-cli doctor",
    "sigmalifting-cli template <kind>",
    'sigmalifting-cli validate <kind> --file <path|"-">',
    'sigmalifting-cli normalize <kind> --file <path|"-">',
    "sigmalifting-cli program <command> [options]",
    "sigmalifting-cli process <command> [options]",
    "sigmalifting-cli rpe <command> [options]",
    "sigmalifting-cli schema <command> [kind]",
    'sigmalifting-cli inspect <view> --file <path|"-">',
    'sigmalifting-cli render <view> --file <path|"-"> [--raw]',
    "sigmalifting-cli xlsx <command> [options]",
    "sigmalifting-cli store <path|init>",
  ],
  commands: {
    root: ROOT_COMMANDS,
    program: PROGRAM_SUBCOMMANDS,
    process: PROCESS_SUBCOMMANDS,
    rpe: RPE_SUBCOMMANDS,
    schema: SCHEMA_SUBCOMMANDS,
    xlsx: XLSX_SUBCOMMANDS,
    store: STORE_SUBCOMMANDS,
  },
  schema_kinds: Object.keys(SCHEMA_CATALOG),
  inspect_views: INSPECT_VIEWS,
  render_views: RENDER_VIEWS,
  flags: {
    file: 'Read a JSON bundle from a file path, file:// URL, "-" for stdin, or an XLSX file for `xlsx import`.',
    data: "Provide an inline JSON payload.",
    data_file: "Read a JSON payload from a file path or file:// URL.",
    out: "Write the resulting bundle or workbook to a file path or file:// URL.",
    exercise_id:
      "Select an exercise for exercise-detail, recording-detail, get-recording, update-set, and update-note.",
    compact: "Print compact JSON instead of pretty JSON.",
    raw: "For render, print raw text directly instead of JSON.",
    store_root:
      "Override the local JSON store root for this invocation. SIGMALIFTING_HOME is the preferred persistent override.",
  },
  text: HELP_TEXT,
});

const handleTemplate = async (kind, flags) => {
  const template = createTemplate(kind, {
    duration: flags.duration ? Number(flags.duration) : undefined,
    weeks: flags.weeks ? Number(flags.weeks) : undefined,
    name: typeof flags.name === "string" ? flags.name : undefined,
    exerciseName:
      typeof flags["exercise-name"] === "string"
        ? flags["exercise-name"]
        : undefined,
    variableParameter:
      typeof flags["variable-parameter"] === "string"
        ? flags["variable-parameter"]
        : undefined,
  });

  return { template };
};

const handleValidate = async (kind, flags, stdinText) => {
  const payload = await readBundle(flags, stdinText);
  const validation = validatePayload(kind, payload);
  return {
    validation,
  };
};

const handleNormalize = async (kind, flags, stdinText) => {
  const payload = await readBundle(flags, stdinText);
  if (kind === "program-import") {
    return normalizeProgramBundle(payload);
  }
  if (kind === "process-import") {
    return normalizeProcessBundle(payload);
  }

  throw new Error(`Normalization is not supported for ${kind}`);
};

const handleProgram = async (subcommand, flags, stdinText, storeOptions) => {
  switch (subcommand) {
    case "list": {
      const listed = await listStoredBundles("program", storeOptions);
      return {
        store: listed.store,
        programs: listed.entries,
        errors: listed.errors,
        persist: false,
      };
    }
    case "show": {
      const programId = getProgramIdFlag(flags);
      if (programId) {
        const stored = await loadStoredBundle("program", programId, storeOptions);
        return {
          bundle: stored.bundle,
          storage: stored.storage,
          persist: false,
        };
      }

      return {
        bundle: await readBundle(flags, stdinText),
        persist: false,
      };
    }
    case "create":
      return createProgramCommand({
        name: typeof flags.name === "string" ? flags.name : undefined,
        description:
          typeof flags.description === "string" ? flags.description : undefined,
      });
    case "copy":
      return copyProgramCommand(await readProgramBundle(flags, stdinText, storeOptions), {
        name: typeof flags.name === "string" ? flags.name : undefined,
      });
    case "update":
      return updateProgramCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        await readPayload(flags, "data", stdinText)
      );
    case "add-block":
      return addBlockCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        await readPayload(flags, "data", stdinText)
      );
    case "update-block":
      return updateBlockCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        await readPayload(flags, "data", stdinText)
      );
    case "delete-block":
      return deleteBlockCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"]
      );
    case "add-day":
      return addDayCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        Number(flags.position),
        await readPayload(flags, "data", stdinText)
      );
    case "rename-day":
      return renameDayCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"],
        flags.name
      );
    case "delete-day":
      return deleteDayCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"]
      );
    case "update-schedule":
      return updateScheduleCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        await readPayload(flags, "data", stdinText)
      );
    case "add-exercise":
      return addExerciseCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"],
        await readPayload(flags, "data", stdinText)
      );
    case "update-exercise":
      return updateExerciseCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["exercise-id"],
        await readPayload(flags, "data", stdinText)
      );
    case "delete-exercise":
      return deleteExerciseCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"],
        flags["exercise-id"]
      );
    case "move-exercise":
      return moveExerciseCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"],
        flags["exercise-id"],
        flags.direction
      );
    case "copy-exercise":
      return copyExerciseCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"],
        flags["exercise-id"]
      );
    case "copy-block":
      return copyBlockCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["random-suffix"] !== "false"
      );
    case "copy-day":
      return copyDayCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["block-id"],
        flags["day-id"]
      );
    case "add-custom-lift":
      return addCustomLiftCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags.name
      );
    case "rename-custom-lift":
      return renameCustomLiftCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags["old-name"],
        flags["new-name"]
      );
    case "delete-custom-lift":
      return deleteCustomLiftCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        flags.name
      );
    default:
      throw new Error(`Unknown program subcommand: ${subcommand}`);
  }
};

const handleProcess = async (subcommand, flags, stdinText, storeOptions) => {
  switch (subcommand) {
    case "list": {
      const listed = await listStoredBundles("process", storeOptions);
      return {
        store: listed.store,
        processes: listed.entries,
        errors: listed.errors,
        persist: false,
      };
    }
    case "show": {
      const processId = getProcessIdFlag(flags);
      if (processId) {
        const stored = await loadStoredBundle("process", processId, storeOptions);
        return {
          bundle: stored.bundle,
          storage: stored.storage,
          persist: false,
        };
      }

      return {
        bundle: await readBundle(flags, stdinText),
        persist: false,
      };
    }
    case "create-from-program":
      return createProcessFromProgramCommand(
        await readProgramBundle(flags, stdinText, storeOptions),
        {
          name: typeof flags.name === "string" ? flags.name : undefined,
          startDate:
            typeof flags["start-date"] === "string"
              ? flags["start-date"]
              : undefined,
          config:
            flags.config || flags["config-file"]
              ? await readPayload(flags, "config", stdinText)
              : {},
          oneRmProfile:
            flags["one-rm"] || flags["one-rm-file"]
              ? await readPayload(flags, "one-rm", stdinText)
              : {},
        }
      );
    case "update":
      return updateProcessCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        await readPayload(flags, "data", stdinText)
      );
    case "update-config":
      return updateProcessConfigCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        await readPayload(flags, "data", stdinText)
      );
    case "update-one-rm":
      return updateProcessOneRmCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        await readPayload(flags, "data", stdinText)
      );
    case "update-config-and-one-rm":
      return updateProcessConfigAndOneRmCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        await readPayload(flags, "one-rm", stdinText),
        await readPayload(flags, "config", stdinText)
      );
    case "update-set":
      return updateProcessSetCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        flags["exercise-id"],
        Number(flags["week-index"]),
        Number(flags["set-index"]),
        await readPayload(flags, "data", stdinText)
      );
    case "update-note":
      return updateProcessNoteCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        flags["exercise-id"],
        Number(flags["week-index"]),
        flags.note
      );
    case "get-recording":
      return getExerciseRecordingCommand(
        await readProcessBundle(flags, stdinText, storeOptions),
        flags["exercise-id"]
      );
    default:
      throw new Error(`Unknown process subcommand: ${subcommand}`);
  }
};

const handleRpe = async (subcommand, flags) => {
  switch (subcommand) {
    case "chart":
      return { chart: RPE_CHART_DATA };
    case "e1rm":
      return {
        estimated_one_rm: calculate1RM(
          Number(flags.weight),
          Number(flags.reps),
          Number(flags.rpe)
        ),
      };
    case "training-weight":
      return {
        training_weight: calculateTrainingWeight(
          Number(flags["one-rm"]),
          Number(flags.reps),
          Number(flags.rpe)
        ),
      };
    case "range":
      return calculateSuggestedWeightRange(
        Number(flags["one-rm"]),
        Number(flags.reps),
        Number(flags.rpe)
      );
    default:
      throw new Error(`Unknown rpe subcommand: ${subcommand}`);
  }
};

const handleSchema = async (subcommand, kind) => {
  if (subcommand === "list") {
    return {
      schemas: Object.values(SCHEMA_CATALOG),
    };
  }

  if (subcommand === "show") {
    if (!kind || !SCHEMA_CATALOG[kind]) {
      throw new Error(`Unknown schema kind: ${kind}`);
    }
    return {
      schema: SCHEMA_CATALOG[kind],
    };
  }

  throw new Error(`Unknown schema subcommand: ${subcommand}`);
};

const handleInspect = async (view, flags, stdinText) => {
  const bundle = await readBundle(flags, stdinText);
  const { bundleKind, page } = buildInspectPage(bundle, view, {
    exerciseId:
      typeof flags["exercise-id"] === "string" ? flags["exercise-id"] : undefined,
  });

  return {
    view,
    bundle_kind: bundleKind,
    page,
  };
};

const handleRender = async (view, flags, stdinText) => {
  const bundle = await readBundle(flags, stdinText);
  const { bundleKind, page } = buildInspectPage(bundle, view, {
    exerciseId:
      typeof flags["exercise-id"] === "string" ? flags["exercise-id"] : undefined,
  });

  return {
    view,
    bundle_kind: bundleKind,
    rendered: renderInspectPage(page),
  };
};

const handleXlsx = async (subcommand, flags, stdinText) => {
  switch (subcommand) {
    case "import":
      return importXlsxFile(flags.file);
    case "export":
      return exportXlsxFile(await readBundle(flags, stdinText), flags.out);
    default:
      throw new Error(`Unknown xlsx subcommand: ${subcommand}`);
  }
};

const handleStore = async (subcommand, storeOptions) => {
  switch (subcommand) {
    case "path":
      return {
        store: getStorePaths(storeOptions),
        persist: false,
      };
    case "init": {
      const initialized = await initStore(storeOptions);
      return {
        store: initialized.store,
        index: initialized.index,
        persist: false,
      };
    }
    default:
      throw new Error(`Unknown store subcommand: ${subcommand}`);
  }
};

export const runCli = async (
  argv,
  {
    stdinText,
    stdout = defaultStdout,
    stderr = defaultStderr,
    storageRoot,
  } = {}
) => {
  const { positionals, flags } = parseArgs(argv);
  const compact = Boolean(flags.compact);
  const commandLabel = positionals.join(" ") || "help";
  const storeOptions = {
    root:
      typeof flags["store-root"] === "string"
        ? flags["store-root"]
        : storageRoot,
  };

  try {
    if (positionals.length === 0 || positionals[0] === "help") {
      stdout(serialize(buildSuccess("help", { help: buildHelpPayload() }), compact));
      return 0;
    }

    let responsePayload;

    switch (positionals[0]) {
      case "template":
        responsePayload = await handleTemplate(positionals[1], flags);
        break;
      case "validate":
        responsePayload = await handleValidate(positionals[1], flags, stdinText);
        break;
      case "normalize":
        responsePayload = await handleNormalize(positionals[1], flags, stdinText);
        break;
      case "program":
        responsePayload = await handleProgram(
          positionals[1],
          flags,
          stdinText,
          storeOptions
        );
        break;
      case "process":
        responsePayload = await handleProcess(
          positionals[1],
          flags,
          stdinText,
          storeOptions
        );
        break;
      case "rpe":
        responsePayload = await handleRpe(positionals[1], flags);
        break;
      case "schema":
        responsePayload = await handleSchema(positionals[1], positionals[2]);
        break;
      case "inspect":
        responsePayload = await handleInspect(positionals[1], flags, stdinText);
        break;
      case "render":
        responsePayload = await handleRender(positionals[1], flags, stdinText);
        break;
      case "xlsx":
        responsePayload = await handleXlsx(positionals[1], flags, stdinText);
        break;
      case "store":
        responsePayload = await handleStore(positionals[1], storeOptions);
        break;
      case "doctor":
        responsePayload = {
          doctor: {
            cwd: process.cwd(),
            supported_commands: ROOT_COMMANDS,
            available_schema_kinds: Object.keys(SCHEMA_CATALOG),
            store: getStorePaths(storeOptions),
          },
        };
        break;
      default:
        throw new Error(`Unknown command: ${positionals[0]}`);
    }

    const shouldPersistBundle = responsePayload.persist !== false;
    const writableBundle =
      responsePayload.bundle || responsePayload.template || null;
    const publicPayload = { ...responsePayload };
    delete publicPayload.persist;

    if (responsePayload.bundle && shouldPersistBundle) {
      publicPayload.storage = await saveBundleToStore(
        responsePayload.bundle,
        storeOptions
      );
    }

    const outputPath =
      writableBundle && flags.out
        ? await writeBundleIfRequested(writableBundle, flags)
        : null;

    if (positionals[0] === "render" && flags.raw) {
      stdout(
        responsePayload.rendered.endsWith("\n")
          ? responsePayload.rendered
          : `${responsePayload.rendered}\n`
      );
      return 0;
    }

    stdout(
      serialize(buildSuccess(commandLabel, publicPayload, outputPath), compact)
    );
    return 0;
  } catch (error) {
    stderr(serialize(buildError(commandLabel, error), compact));
    return 1;
  }
};
