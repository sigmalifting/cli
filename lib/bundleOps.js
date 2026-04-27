import { generateId } from "./idGenerator.js";
import {
  safeSanitizeDescription,
  safeSanitizeName,
  safeSanitizeNote,
  validateCustomLiftName,
  formatValidationErrors,
} from "./validation/shared.js";
import {
  validateProgramImport,
  programSchema,
} from "./validation/programSchemas.js";
import {
  validateProcessImport,
  processSchema,
} from "./validation/processSchemas.js";
import {
  validateExercise,
  exerciseSchema,
} from "./validation/exerciseSchemas.js";
import { validateImportBusinessLogic } from "./validation/businessLogicSchemas.js";

const DEFAULT_APP_NAME = "SigmaLifting";
const DEFAULT_VERSION = "cli";
const EMPTY_WEEKLY_SCHEDULE = ["", "", "", "", "", "", ""];
const STANDARD_ONE_RM_LIFTS = ["squat", "bench", "deadlift"];
const PROFILE_CONTROL_KEYS = ["enable_blockly_one_rm", "blockly_one_rm"];
const PROGRAM_UPDATE_FIELDS = ["name", "description"];
const ADD_BLOCK_FIELDS = ["name", "duration"];
const UPDATE_BLOCK_FIELDS = ["name", "duration"];
const ADD_DAY_FIELDS = ["name"];
const ADD_EXERCISE_FIELDS = [
  "exercise_name",
  "name",
  "variable_parameter",
  "variableParameter",
  "set_groups",
  "one_rm_anchor",
  "deload_config",
];
const UPDATE_EXERCISE_FIELDS = [
  "exercise_name",
  "one_rm_anchor",
  "deload_config",
];
const PROCESS_UPDATE_FIELDS = ["name"];
const PROCESS_CONFIG_FIELDS = ["weight_unit", "weight_rounding", "week_start_day"];
const PROCESS_SET_FIELDS = ["weight", "reps", "rpe", "completed"];
const VARIABLE_PARAMETERS = ["reps", "weight", "rpe"];
const WEIGHT_MODEL_MODES = ["percentage", "mixed"];
const DYNAMIC_WEIGHT_TYPES = ["%_of_weight", "target_rpe", ""];
const VARIABLE_PARAMETER_FIELDS = {
  reps: "weekly_reps",
  rpe: "weekly_rpe",
  weight: "weekly_weight_percentage",
};
const WEEK_VALUE_FIELDS = [
  "weekly_num_sets",
  "weekly_reps",
  "weekly_rpe",
  "weekly_weight_percentage",
  "weekly_notes",
  "mix_weight_config",
  "backoff_%",
  "backoff_rpe",
  "drop_%",
  "drop_rpe",
  "rpe_cap",
];
const MIX_WEIGHT_SUB_FIELDS = [
  "weekly_weight_percentage",
  "weekly_weight_absolute",
  "weight_unit",
];

const deepClone = (value) => JSON.parse(JSON.stringify(value));
const nowIso = () => new Date().toISOString();
const nowDate = () => new Date().toISOString().slice(0, 10);
const randomSuffix = () => generateId("").replace(/-/g, "").slice(0, 3);

const createCliError = (message, details) => {
  const error = new Error(message);
  if (details !== undefined) {
    error.details = details;
  }
  return error;
};

const ensure = (condition, message, details) => {
  if (!condition) {
    throw createCliError(message, details);
  }
};

const ensurePlainObject = (value, label) => {
  ensure(
    value !== null && typeof value === "object" && !Array.isArray(value),
    `${label} must be a JSON object`
  );
};

const pickAllowedPatch = (patch, allowedFields, label) => {
  ensurePlainObject(patch, label);
  const allowed = new Set(allowedFields);
  const unsupported = Object.keys(patch).filter((field) => !allowed.has(field));
  ensure(
    unsupported.length === 0,
    `${label} contains unsupported field(s): ${unsupported.join(", ")}. Supported fields: ${allowedFields.join(", ")}`
  );
  return patch;
};

const normalizeLiftKey = (liftName) =>
  typeof liftName === "string" ? liftName.trim().toLowerCase() : "";

const getAllowedOneRmLiftKeys = (program) =>
  new Set([
    ...STANDARD_ONE_RM_LIFTS,
    ...((program.custom_anchored_lifts || []).map(normalizeLiftKey).filter(Boolean)),
  ]);

const validateOneRmProfileKeys = (program, profile, label) => {
  ensurePlainObject(profile, label);

  const allowedLiftKeys = getAllowedOneRmLiftKeys(program);
  const allowedTopLevelKeys = new Set([
    ...allowedLiftKeys,
    ...PROFILE_CONTROL_KEYS,
  ]);
  const unsupportedTopLevelKeys = Object.keys(profile).filter(
    (key) => !allowedTopLevelKeys.has(key)
  );

  ensure(
    unsupportedTopLevelKeys.length === 0,
    `${label} contains unsupported lift key(s): ${unsupportedTopLevelKeys.join(", ")}. Add the custom lift to the program first.`
  );

  if (profile.blockly_one_rm !== undefined) {
    ensure(
      Array.isArray(profile.blockly_one_rm),
      `${label}.blockly_one_rm must be an array`
    );

    profile.blockly_one_rm.forEach((blockProfile, blockIndex) => {
      ensurePlainObject(
        blockProfile,
        `${label}.blockly_one_rm[${blockIndex}]`
      );
      const unsupportedBlockKeys = Object.keys(blockProfile).filter(
        (key) => !allowedLiftKeys.has(key)
      );
      ensure(
        unsupportedBlockKeys.length === 0,
        `${label}.blockly_one_rm[${blockIndex}] contains unsupported lift key(s): ${unsupportedBlockKeys.join(", ")}. Add the custom lift to the program first.`
      );
    });
  }
};

const mergeOneRmProfile = (existingProfile = {}, patch = {}) => {
  const merged = {
    ...deepClone(existingProfile),
  };

  Object.entries(patch).forEach(([key, value]) => {
    if (key !== "blockly_one_rm") {
      merged[key] = value;
    }
  });

  if (patch.blockly_one_rm !== undefined) {
    const mergedBlocks = Array.isArray(merged.blockly_one_rm)
      ? merged.blockly_one_rm.map((entry) => ({ ...entry }))
      : [];

    patch.blockly_one_rm.forEach((blockPatch, blockIndex) => {
      mergedBlocks[blockIndex] = {
        ...(mergedBlocks[blockIndex] || {}),
        ...blockPatch,
      };
    });

    merged.blockly_one_rm = mergedBlocks;
  }

  return merged;
};

const resizeArray = (array, targetLength, fallbackValue, isSimpleDeload = false) => {
  if (!Array.isArray(array)) {
    return array;
  }

  if (array.length > targetLength) {
    return array.slice(0, targetLength);
  }

  const resized = [...array];
  if (
    isSimpleDeload &&
    resized.length > 1 &&
    resized[resized.length - 1] === -1
  ) {
    resized[resized.length - 1] = resized[resized.length - 2];
  }

  const fillValue = resized.length > 0 ? resized[resized.length - 1] : fallbackValue;
  while (resized.length < targetLength) {
    resized.push(fillValue);
  }

  return resized;
};

const resizeNamedArray = (
  target,
  key,
  targetLength,
  fallbackValue,
  isSimpleDeload = false
) => {
  if (target && Array.isArray(target[key])) {
    target[key] = resizeArray(
      target[key],
      targetLength,
      fallbackValue,
      isSimpleDeload
    );
  }
};

const resizeExerciseWeeklyArraysForBlock = (bundle, blockId, duration) => {
  bundle.exercises
    .filter((exercise) => exercise.block_id === blockId)
    .forEach((exercise) => {
      const isSimpleDeload = exercise.deload_config?.enabled === true;
      exercise.set_groups?.forEach((group) => {
        resizeNamedArray(group, "weekly_notes", duration, "", isSimpleDeload);
        resizeNamedArray(group, "weekly_num_sets", duration, 3, isSimpleDeload);
        resizeNamedArray(group, "weekly_reps", duration, 5, isSimpleDeload);
        resizeNamedArray(group, "weekly_rpe", duration, 8, isSimpleDeload);
        resizeNamedArray(
          group,
          "weekly_weight_percentage",
          duration,
          65,
          isSimpleDeload
        );

        resizeNamedArray(
          group.mix_weight_config,
          "weekly_weight_percentage",
          duration,
          70,
          isSimpleDeload
        );
        resizeNamedArray(
          group.mix_weight_config,
          "weekly_weight_absolute",
          duration,
          0,
          isSimpleDeload
        );
        resizeNamedArray(
          group.backoff_config,
          "weekly_rpe",
          duration,
          8,
          isSimpleDeload
        );
        resizeNamedArray(
          group.backoff_config,
          "weekly_percentage",
          duration,
          80,
          isSimpleDeload
        );
        resizeNamedArray(
          group.fatigue_drop_config,
          "rpe_cap",
          duration,
          8.5,
          isSimpleDeload
        );
        resizeNamedArray(
          group.fatigue_drop_config,
          "weekly_rpe",
          duration,
          8,
          isSimpleDeload
        );
        resizeNamedArray(
          group.fatigue_drop_config,
          "weekly_percentage",
          duration,
          80,
          isSimpleDeload
        );
      });
    });
};

const getExerciseBlockDuration = (bundle, exerciseId) => {
  const { exercise } = findExercise(bundle, exerciseId);
  const { block } = findBlock(bundle, exercise.block_id);
  return block.duration;
};

const getPrescribedSetCount = (bundle, exerciseId, weekIndex) => {
  const { exercise } = findExercise(bundle, exerciseId);
  return (exercise.set_groups || []).reduce(
    (total, group) => total + (group.weekly_num_sets?.[weekIndex] || 0),
    0
  );
};

const findSetGroup = (exercise, groupId) => {
  const setGroups = Array.isArray(exercise.set_groups)
    ? exercise.set_groups
    : [];
  const index = setGroups.findIndex((group) => group.group_id === groupId);
  ensure(index !== -1, `Set group ${groupId} not found`);
  return {
    group: setGroups[index],
    groupIndex: index,
  };
};

const getExerciseDuration = (bundle, exercise) => {
  const { block } = findBlock(bundle, exercise.block_id);
  return block.duration;
};

const ensureWeekIndexInExercise = (bundle, exercise, weekIndex) => {
  ensure(
    Number.isInteger(weekIndex) && weekIndex >= 0,
    "week-index must be a non-negative integer"
  );
  const duration = getExerciseDuration(bundle, exercise);
  ensure(
    weekIndex < duration,
    `week-index ${weekIndex} is outside the exercise block duration (${duration})`
  );
  return duration;
};

const ensureArrayValueSlot = (target, key, duration, fallbackValue) => {
  if (!Array.isArray(target[key])) {
    target[key] = Array(duration).fill(fallbackValue);
    return;
  }

  target[key] = resizeArray(target[key], duration, fallbackValue);
};

const setWeeklyArrayValue = (target, key, weekIndex, value, duration, fallback) => {
  ensureArrayValueSlot(target, key, duration, fallback);
  target[key][weekIndex] = value;
};

const getDefaultForWeekValueField = (field, subField) => {
  if (field === "weekly_num_sets") return 3;
  if (field === "weekly_reps") return 5;
  if (field === "weekly_rpe") return 8;
  if (field === "weekly_weight_percentage") return 65;
  if (field === "weekly_notes") return "";
  if (field === "backoff_%") return 80;
  if (field === "backoff_rpe") return 8;
  if (field === "drop_%") return 80;
  if (field === "drop_rpe") return 8;
  if (field === "rpe_cap") return 8.5;
  if (field === "mix_weight_config" && subField === "weekly_weight_absolute") {
    return 0;
  }
  return 70;
};

const parseNumericFlagValue = (value, label) => {
  ensure(
    typeof value === "number" || typeof value === "string",
    `${label} requires a numeric value`
  );
  ensure(
    typeof value !== "string" || value.trim() !== "",
    `${label} requires a numeric value`
  );
  const numeric = Number(value);
  ensure(Number.isFinite(numeric), `${label} must be a finite number`);
  return numeric;
};

const normalizeDynamicWeightType = (type) =>
  typeof type === "string" ? type.trim().toLowerCase() : "";

const requireRpeVariableParameter = (group, featureLabel) => {
  ensure(
    group.variable_parameter === "rpe",
    `${featureLabel} requires variable_parameter to be 'rpe'`
  );
};

const clearVariableParameterField = (group, variableParameter, duration) => {
  const field = VARIABLE_PARAMETER_FIELDS[variableParameter];
  if (field) {
    group[field] = Array(duration).fill(-1);
  }
};

const ensureWeekValueFieldIsEditable = (group, field) => {
  const variableField = VARIABLE_PARAMETER_FIELDS[group.variable_parameter];
  ensure(
    field !== variableField,
    `Cannot set ${field} when ${group.variable_parameter} is the variable parameter`
  );
  ensure(
    field !== "mix_weight_config" || group.variable_parameter !== "weight",
    "Cannot set mix_weight_config when weight is the variable parameter"
  );
};

const touchMetadata = (bundle) => ({
  ...bundle,
  metadata: {
    exportedAt: nowIso(),
    version: bundle.metadata?.version || DEFAULT_VERSION,
    appName: bundle.metadata?.appName || DEFAULT_APP_NAME,
  },
});

const touchProgramTimestamp = (bundle) => {
  bundle.program.updated_at = nowIso();
};

const touchProcessTimestamp = (bundle) => {
  bundle.process.updated_at = nowIso();
};

const buildDefaultSetGroup = ({
  weeks = 4,
  variableParameter = "weight",
} = {}) => {
  const weeklyReps =
    variableParameter === "reps"
      ? Array(weeks).fill(-1)
      : Array(weeks).fill(5);
  const weeklyRpe =
    variableParameter === "rpe"
      ? Array(weeks).fill(-1)
      : Array(weeks).fill(8);
  const weeklyWeightPercentage =
    variableParameter === "weight"
      ? Array(weeks).fill(-1)
      : Array(weeks).fill(65);

  return {
    group_id: generateId("sg"),
    variable_parameter: variableParameter,
    weekly_notes: Array(weeks).fill(""),
    weekly_num_sets: Array(weeks).fill(3),
    weekly_reps: weeklyReps,
    weekly_rpe: weeklyRpe,
    weekly_weight_percentage: weeklyWeightPercentage,
    mix_weight_config: {
      enabled: false,
      weight_unit: "kg",
    },
    backoff_config: {
      enabled: false,
      depends_on_set_group_id: "",
      type: "",
      weekly_percentage: Array(weeks).fill(80),
      weekly_rpe: Array(weeks).fill(8),
    },
    fatigue_drop_config: {
      enabled: false,
      type: "",
      weekly_percentage: Array(weeks).fill(80),
      weekly_rpe: Array(weeks).fill(8),
      rpe_cap: Array(weeks).fill(8.5),
    },
  };
};

const buildExerciseDefinition = ({
  programId,
  blockId,
  exerciseName = "New Exercise",
  weeks = 4,
  variableParameter = "weight",
  setGroups,
  oneRmAnchor,
  deloadConfig,
} = {}) => ({
  _id: generateId("ex"),
  program_id: programId,
  block_id: blockId,
  exercise_name: safeSanitizeName(exerciseName),
  one_rm_anchor: oneRmAnchor || { enabled: false },
  deload_config: deloadConfig || {
    enabled: false,
    percentage: 85,
  },
  set_groups:
    setGroups && setGroups.length > 0
      ? setGroups
      : [buildDefaultSetGroup({ weeks, variableParameter })],
  created_at: nowIso(),
  updated_at: nowIso(),
});

const buildTrainingDay = ({ name = "New Day", exercises = [] } = {}) => ({
  _id: generateId("day"),
  name: safeSanitizeName(name),
  exercises,
});

const buildBlock = ({
  programId,
  name = "New Block",
  duration = 4,
  weeklySchedule,
  trainingDays,
} = {}) => ({
  _id: generateId("block"),
  name: safeSanitizeName(name),
  program_id: programId,
  duration,
  weekly_schedule: weeklySchedule || [...EMPTY_WEEKLY_SCHEDULE],
  training_days: trainingDays || [],
});

const buildProcessRecording = (exerciseId, weeks) => ({
  exercise_id: exerciseId,
  weekly: Array.from({ length: weeks }, () => ({
    sets: [],
    note: "",
  })),
});

const parseSchema = (schema, data) => {
  try {
    return { success: true, data: schema.parse(data) };
  } catch (error) {
    return {
      success: false,
      error: formatValidationErrors(error),
      details: error.format ? error.format() : null,
    };
  }
};

const normalizeBundleResult = (
  validation,
  kind,
  { rejectWarnings = false } = {}
) => {
  if (!validation.success) {
    throw createCliError(`Invalid ${kind} payload`, validation);
  }

  const businessValidation = validateImportBusinessLogic(validation.data, kind);
  if (!businessValidation.success) {
    throw createCliError(`Invalid ${kind} business logic`, businessValidation);
  }

  const warnings = businessValidation.warnings || [];
  if (rejectWarnings && warnings.length > 0) {
    const warningSummary = warnings
      .slice(0, 3)
      .map((warning) => warning.code || warning.message)
      .join(", ");
    throw createCliError(`${kind} mutation would require cleanup: ${warningSummary}`, {
      success: false,
      warnings,
    });
  }

  return {
    bundle: touchMetadata(businessValidation.data),
    warnings,
  };
};

const validateBundleResultStrict = (validation, kind) => {
  if (!validation.success) {
    return validation;
  }

  const businessValidation = validateImportBusinessLogic(validation.data, kind);
  const warnings = businessValidation.warnings || [];

  if (businessValidation.success && warnings.length > 0) {
    return {
      success: false,
      errors: warnings,
      warnings,
      data: businessValidation.data,
    };
  }

  return businessValidation;
};

const remapSetGroups = (exercise, setGroupIdMap) => {
  if (!Array.isArray(exercise.set_groups)) {
    return;
  }

  exercise.set_groups.forEach((setGroup) => {
    const oldGroupId = setGroup.group_id;
    const newGroupId = generateId("sg");
    setGroupIdMap[oldGroupId] = newGroupId;
    setGroup.group_id = newGroupId;
  });

  exercise.set_groups.forEach((setGroup) => {
    const dependency = setGroup.backoff_config?.depends_on_set_group_id;
    if (dependency && setGroupIdMap[dependency]) {
      setGroup.backoff_config.depends_on_set_group_id = setGroupIdMap[dependency];
    }
  });
};

const findBlock = (bundle, blockId) => {
  const index = bundle.program.blocks.findIndex((block) => block._id === blockId);
  ensure(index !== -1, `Block ${blockId} not found`);
  return {
    block: bundle.program.blocks[index],
    blockIndex: index,
  };
};

const findDay = (block, dayId) => {
  const index = block.training_days.findIndex((day) => day._id === dayId);
  ensure(index !== -1, `Training day ${dayId} not found`);
  return {
    day: block.training_days[index],
    dayIndex: index,
  };
};

const findExercise = (bundle, exerciseId) => {
  const index = bundle.exercises.findIndex((exercise) => exercise._id === exerciseId);
  ensure(index !== -1, `Exercise ${exerciseId} not found`);
  return {
    exercise: bundle.exercises[index],
    exerciseIndex: index,
  };
};

const getOrderedTrainingDayIds = (block) =>
  (block.weekly_schedule || []).filter((dayId) => dayId && dayId !== "");

const applyProgramMutation = (bundle, mutateFn) => {
  const working = deepClone(bundle);
  const mutationResult = mutateFn(working);
  touchProgramTimestamp(working);
  const normalized = normalizeBundleResult(
    validateProgramImport(touchMetadata(working)),
    "program",
    { rejectWarnings: true }
  );

  return {
    bundle: normalized.bundle,
    warnings: normalized.warnings,
    result: mutationResult,
  };
};

const applyProcessMutation = (bundle, mutateFn) => {
  const working = deepClone(bundle);
  const mutationResult = mutateFn(working);
  touchProcessTimestamp(working);
  const normalized = normalizeBundleResult(
    validateProcessImport(touchMetadata(working)),
    "process",
    { rejectWarnings: true }
  );

  return {
    bundle: normalized.bundle,
    warnings: normalized.warnings,
    result: mutationResult,
  };
};

export const SCHEMA_CATALOG = {
  "program-import": {
    kind: "program-import",
    description:
      "Canonical app import/export bundle: { program, exercises, metadata? }",
    required_keys: ["program", "exercises"],
    notes: [
      "Use this for constructing mobile-compatible program JSON.",
      "Validate rejects bundles that require cleanup; normalize is the explicit repair path.",
      "When one_rm_anchor.enabled is true, lift_type must be squat, bench, deadlift, or a declared custom_anchored_lifts entry.",
      "Each set group records exactly one variable parameter; backoff and fatigue-drop configs are only valid when variable_parameter is rpe.",
    ],
  },
  "process-import": {
    kind: "process-import",
    description:
      "Canonical app import/export bundle: { process, program, exercises, metadata? }",
    required_keys: ["process", "program", "exercises"],
    notes: [
      "This is the exact JSON shape the app imports for a process snapshot.",
      "Process recordings must align with exercise ids from the embedded program snapshot.",
      "Embedded program exercises must satisfy program-import anchor rules.",
    ],
  },
  program: {
    kind: "program",
    description: "Program entity only",
    required_keys: ["_id", "name", "created_at", "updated_at", "blocks"],
  },
  process: {
    kind: "process",
    description: "Process entity only",
    required_keys: [
      "_id",
      "name",
      "program_id",
      "program_name",
      "created_at",
      "updated_at",
      "start_date",
      "config",
      "one_rm_profile",
    ],
  },
  exercise: {
    kind: "exercise",
    description: "Exercise entity only",
    required_keys: [
      "_id",
      "program_id",
      "block_id",
      "exercise_name",
      "one_rm_anchor",
      "set_groups",
    ],
    notes: [
      "If one_rm_anchor.enabled is true, lift_type is required. Full allowed-lift checking needs a program-import or process-import bundle.",
    ],
  },
};

export const createProgramBundle = ({ name = "New Program", description = "" } = {}) => {
  const timestamp = nowIso();

  return normalizeBundleResult(
    validateProgramImport({
      program: {
        _id: generateId("prog"),
        name: safeSanitizeName(name),
        description: safeSanitizeDescription(description || ""),
        created_at: timestamp,
        updated_at: timestamp,
        blocks: [],
        custom_anchored_lifts: [],
      },
      exercises: [],
      metadata: {
        exportedAt: timestamp,
        version: DEFAULT_VERSION,
        appName: DEFAULT_APP_NAME,
      },
    }),
    "program"
  );
};

export const createExampleProgramBundle = ({
  name = "Example Program",
  description = "Valid starter bundle for SigmaLifting agents",
  duration = 4,
  dayPosition = 1,
  exerciseName = "Competition Squat",
} = {}) => {
  const base = createProgramBundle({ name, description }).bundle;
  const block = buildBlock({
    programId: base.program._id,
    name: "Block 1",
    duration,
  });
  const day = buildTrainingDay({ name: "Day 1" });
  block.training_days.push(day);
  block.weekly_schedule[dayPosition] = day._id;

  const exercise = buildExerciseDefinition({
    programId: base.program._id,
    blockId: block._id,
    exerciseName,
    weeks: duration,
  });
  day.exercises.push(exercise._id);

  base.program.blocks.push(block);
  base.exercises.push(exercise);

  return normalizeBundleResult(validateProgramImport(touchMetadata(base)), "program");
};

export const createProcessBundleFromProgram = (
  programBundle,
  {
    name,
    startDate = nowDate(),
    config = {},
    oneRmProfile = {},
  } = {}
) => {
  const normalizedProgram = normalizeProgramBundle(programBundle).bundle;
  const program = deepClone(normalizedProgram.program);
  const exercises = deepClone(normalizedProgram.exercises);
  const timestamp = nowIso();
  const customLifts = {};
  const configPatch = pickAllowedPatch(
    config,
    PROCESS_CONFIG_FIELDS,
    "Process config"
  );

  (program.custom_anchored_lifts || []).forEach((liftName) => {
    customLifts[liftName.toLowerCase()] = -1;
  });

  validateOneRmProfileKeys(program, oneRmProfile, "One RM profile");

  const defaultOneRmProfile = {
    enable_blockly_one_rm: false,
    squat: -1,
    bench: -1,
    deadlift: -1,
    ...customLifts,
    blockly_one_rm: program.blocks.map(() => ({
      squat: -1,
      bench: -1,
      deadlift: -1,
      ...customLifts,
    })),
  };

  const process = {
    _id: generateId("proc"),
    name: safeSanitizeName(name || `${program.name} Process`),
    program_id: program._id,
    program_name: program.name,
    created_at: timestamp,
    updated_at: timestamp,
    start_date: startDate,
    config: {
      weight_unit: "kg",
      week_start_day: 0,
      ...configPatch,
    },
    one_rm_profile: mergeOneRmProfile(defaultOneRmProfile, oneRmProfile),
    exercise_recordings: exercises.map((exercise) => {
      const block = program.blocks.find((candidate) => candidate._id === exercise.block_id);
      const weeks = block?.duration || 1;
      return buildProcessRecording(exercise._id, weeks);
    }),
  };

  return normalizeBundleResult(
    validateProcessImport(
      touchMetadata({
        process,
        program,
        exercises,
      })
    ),
    "process"
  );
};

export const createExampleProcessBundle = (options = {}) =>
  createProcessBundleFromProgram(createExampleProgramBundle(options).bundle, options);

export const createTemplate = (kind, options = {}) => {
  switch (kind) {
    case "program-import":
      return createExampleProgramBundle(options).bundle;
    case "process-import":
      return createExampleProcessBundle(options).bundle;
    case "set-group":
      return buildDefaultSetGroup(options);
    case "training-day":
      return buildTrainingDay(options);
    case "block":
      return buildBlock({
        programId: options.programId || "prog_TEMPLATE",
        name: options.name,
        duration: options.duration || 4,
      });
    case "exercise":
      return buildExerciseDefinition({
        programId: options.programId || "prog_TEMPLATE",
        blockId: options.blockId || "block_TEMPLATE",
        exerciseName: options.exerciseName || "Template Exercise",
        weeks: options.weeks || 4,
        variableParameter: options.variableParameter || "weight",
      });
    case "process-config":
      return {
        weight_unit: "kg",
        weight_rounding: 2.5,
        week_start_day: 0,
      };
    case "one-rm-profile":
      return {
        enable_blockly_one_rm: false,
        squat: -1,
        bench: -1,
        deadlift: -1,
        blockly_one_rm: [{ squat: -1, bench: -1, deadlift: -1 }],
      };
    case "exercise-recording":
      return buildProcessRecording(
        options.exerciseId || "ex_TEMPLATE",
        options.weeks || 4
      );
    default:
      throw createCliError(`Unknown template kind: ${kind}`);
  }
};

export const validatePayload = (kind, data) => {
  switch (kind) {
    case "program-import": {
      const validation = validateProgramImport(data);
      return validateBundleResultStrict(validation, "program");
    }
    case "process-import": {
      const validation = validateProcessImport(data);
      return validateBundleResultStrict(validation, "process");
    }
    case "program":
      return parseSchema(programSchema, data);
    case "process":
      return parseSchema(processSchema, data);
    case "exercise":
      return validateExercise(data);
    default:
      throw createCliError(`Unknown validation kind: ${kind}`);
  }
};

export const normalizeProgramBundle = (bundle) =>
  normalizeBundleResult(validateProgramImport(bundle), "program");

export const normalizeProcessBundle = (bundle) =>
  normalizeBundleResult(validateProcessImport(bundle), "process");

export const createProgramCommand = (options = {}) => createProgramBundle(options);

export const copyProgramCommand = (bundle, options = {}) => {
  const normalized = normalizeProgramBundle(bundle).bundle;
  const working = deepClone(normalized);
  const originalProgramId = working.program._id;
  const newProgramId = generateId("prog");
  const idMap = {
    blocks: {},
    trainingDays: {},
    exercises: {},
    setGroups: {},
  };
  const timestamp = nowIso();

  working.program._id = newProgramId;
  working.program.name = safeSanitizeName(
    options.name || `${working.program.name} _copy ${randomSuffix()}`
  );
  working.program.created_at = timestamp;
  working.program.updated_at = timestamp;

  working.program.blocks.forEach((block) => {
    const oldBlockId = block._id;
    const newBlockId = generateId("block");
    idMap.blocks[oldBlockId] = newBlockId;
    block._id = newBlockId;
    block.program_id = newProgramId;

    block.training_days.forEach((day) => {
      const oldDayId = day._id;
      const newDayId = generateId("day");
      idMap.trainingDays[oldDayId] = newDayId;
      day._id = newDayId;
    });

    block.weekly_schedule = block.weekly_schedule.map((dayId) =>
      dayId && idMap.trainingDays[dayId] ? idMap.trainingDays[dayId] : dayId
    );
  });

  working.exercises.forEach((exercise) => {
    const oldExerciseId = exercise._id;
    const newExerciseId = generateId("ex");
    idMap.exercises[oldExerciseId] = newExerciseId;
    exercise._id = newExerciseId;
    exercise.program_id = newProgramId;
    exercise.block_id = idMap.blocks[exercise.block_id];
    exercise.created_at = timestamp;
    exercise.updated_at = timestamp;
    remapSetGroups(exercise, idMap.setGroups);
  });

  working.program.blocks.forEach((block) => {
    block.training_days.forEach((day) => {
      day.exercises = day.exercises.map((exerciseId) => idMap.exercises[exerciseId] || exerciseId);
    });
  });

  return normalizeProgramBundle(touchMetadata(working));
};

export const updateProgramCommand = (bundle, patch) =>
  applyProgramMutation(bundle, (working) => {
    const safePatch = pickAllowedPatch(
      patch,
      PROGRAM_UPDATE_FIELDS,
      "Program update payload"
    );
    working.program = {
      ...working.program,
      ...safePatch,
      name: safePatch.name ? safeSanitizeName(safePatch.name) : working.program.name,
      description:
        safePatch.description !== undefined
          ? safeSanitizeDescription(safePatch.description)
          : working.program.description,
    };

    return {
      program_id: working.program._id,
    };
  });

export const addBlockCommand = (bundle, blockData = {}) =>
  applyProgramMutation(bundle, (working) => {
    const safeBlockData = pickAllowedPatch(
      blockData,
      ADD_BLOCK_FIELDS,
      "Block create payload"
    );
    const block = buildBlock({
      programId: working.program._id,
      name: safeBlockData.name || "New Block",
      duration: safeBlockData.duration || 4,
      weeklySchedule: [...EMPTY_WEEKLY_SCHEDULE],
      trainingDays: [],
    });

    working.program.blocks.push(block);

    return {
      block_id: block._id,
    };
  });

export const updateBlockCommand = (bundle, blockId, blockPatch = {}) =>
  applyProgramMutation(bundle, (working) => {
    const safeBlockPatch = pickAllowedPatch(
      blockPatch,
      UPDATE_BLOCK_FIELDS,
      "Block update payload"
    );
    const { block, blockIndex } = findBlock(working, blockId);
    const nextDuration =
      safeBlockPatch.duration !== undefined ? safeBlockPatch.duration : block.duration;
    working.program.blocks[blockIndex] = {
      ...block,
      ...safeBlockPatch,
      name: safeBlockPatch.name ? safeSanitizeName(safeBlockPatch.name) : block.name,
    };
    resizeExerciseWeeklyArraysForBlock(working, blockId, nextDuration);

    return {
      block_id: blockId,
    };
  });

export const deleteBlockCommand = (bundle, blockId) =>
  applyProgramMutation(bundle, (working) => {
    const { blockIndex } = findBlock(working, blockId);
    working.program.blocks.splice(blockIndex, 1);
    working.exercises = working.exercises.filter((exercise) => exercise.block_id !== blockId);

    return {
      block_id: blockId,
    };
  });

export const addDayCommand = (bundle, blockId, position, dayData = {}) =>
  applyProgramMutation(bundle, (working) => {
    const safeDayData = pickAllowedPatch(
      dayData,
      ADD_DAY_FIELDS,
      "Training day create payload"
    );
    ensure(Number.isInteger(position) && position >= 0 && position <= 6, "Day position must be between 0 and 6");
    const { block } = findBlock(working, blockId);
    ensure(!block.weekly_schedule[position], `Weekly schedule already has a day at position ${position}`);

    const day = buildTrainingDay({
      name: safeDayData.name || "New Day",
      exercises: [],
    });
    block.training_days.push(day);
    block.weekly_schedule[position] = day._id;

    return {
      day_id: day._id,
      position,
    };
  });

export const renameDayCommand = (bundle, blockId, dayId, name) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);
    day.name = safeSanitizeName(name);

    return {
      day_id: dayId,
    };
  });

export const deleteDayCommand = (bundle, blockId, dayId) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);
    const removedExerciseIds = new Set(day.exercises || []);
    block.training_days = block.training_days.filter((day) => day._id !== dayId);
    block.weekly_schedule = block.weekly_schedule.map((entry) =>
      entry === dayId ? "" : entry
    );
    const remainingExerciseIds = new Set(
      working.program.blocks.flatMap((candidateBlock) =>
        candidateBlock.training_days.flatMap((candidateDay) => candidateDay.exercises || [])
      )
    );
    working.exercises = working.exercises.filter(
      (exercise) =>
        !removedExerciseIds.has(exercise._id) || remainingExerciseIds.has(exercise._id)
    );

    return {
      day_id: dayId,
    };
  });

export const updateScheduleCommand = (bundle, blockId, weeklySchedule) =>
  applyProgramMutation(bundle, (working) => {
    ensure(Array.isArray(weeklySchedule) && weeklySchedule.length === 7, "Weekly schedule must be an array of 7 entries");
    const { block } = findBlock(working, blockId);
    const trainingDayIds = new Set(block.training_days.map((day) => day._id));
    const invalidDayIds = weeklySchedule.filter(
      (dayId) => dayId && !trainingDayIds.has(dayId)
    );
    ensure(
      invalidDayIds.length === 0,
      `Weekly schedule references unknown training day ID(s): ${invalidDayIds.join(", ")}`
    );
    block.weekly_schedule = weeklySchedule;

    return {
      block_id: blockId,
    };
  });

export const addExerciseCommand = (bundle, blockId, dayId, exerciseData = {}) =>
  applyProgramMutation(bundle, (working) => {
    const safeExerciseData = pickAllowedPatch(
      exerciseData,
      ADD_EXERCISE_FIELDS,
      "Exercise create payload"
    );
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);
    const exercise = buildExerciseDefinition({
      programId: working.program._id,
      blockId,
      exerciseName:
        safeExerciseData.exercise_name || safeExerciseData.name || "New Exercise",
      weeks: block.duration,
      variableParameter:
        safeExerciseData.variable_parameter ||
        safeExerciseData.variableParameter ||
        "weight",
      setGroups: safeExerciseData.set_groups,
      oneRmAnchor: safeExerciseData.one_rm_anchor,
      deloadConfig: safeExerciseData.deload_config,
    });

    working.exercises.push(exercise);
    resizeExerciseWeeklyArraysForBlock(working, blockId, block.duration);
    day.exercises.push(exercise._id);

    return {
      exercise_id: exercise._id,
    };
  });

export const updateExerciseCommand = (bundle, exerciseId, patch = {}) =>
  applyProgramMutation(bundle, (working) => {
    const safePatch = pickAllowedPatch(
      patch,
      UPDATE_EXERCISE_FIELDS,
      "Exercise update payload"
    );
    const { exercise } = findExercise(working, exerciseId);
    Object.assign(exercise, safePatch);
    if (safePatch.exercise_name) {
      exercise.exercise_name = safeSanitizeName(safePatch.exercise_name);
    }
    exercise.updated_at = nowIso();
    const { block } = findBlock(working, exercise.block_id);
    resizeExerciseWeeklyArraysForBlock(working, exercise.block_id, block.duration);

    const validation = validateExercise(exercise);
    if (!validation.success) {
      throw createCliError(`Updated exercise ${exerciseId} is invalid`, validation);
    }

    return {
      exercise_id: exerciseId,
    };
  });

export const setExerciseAnchorCommand = (
  bundle,
  exerciseId,
  { enabled, liftType, ratio } = {}
) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    const isEnabled = Boolean(enabled);

    if (!isEnabled) {
      exercise.one_rm_anchor = { enabled: false };
    } else {
      ensure(
        typeof liftType === "string" && liftType.trim().length > 0,
        "Enabled anchor requires --lift-type"
      );
      exercise.one_rm_anchor = {
        enabled: true,
        lift_type: liftType.trim(),
        ratio:
          ratio !== undefined
            ? parseNumericFlagValue(ratio, "Anchor ratio")
            : exercise.one_rm_anchor?.ratio || 0.8,
      };
    }

    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
    };
  });

export const setExerciseDeloadCommand = (
  bundle,
  exerciseId,
  { enabled, percentage } = {}
) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    exercise.deload_config = {
      enabled: Boolean(enabled),
      percentage:
        percentage !== undefined
          ? parseNumericFlagValue(percentage, "Deload percentage")
          : exercise.deload_config?.percentage || 85,
    };
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
    };
  });

export const addSetGroupCommand = (bundle, exerciseId) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    ensure(exercise.set_groups.length < 15, "At most 15 set groups");
    const duration = getExerciseDuration(working, exercise);
    const setGroup = buildDefaultSetGroup({ weeks: duration });
    exercise.set_groups.push(setGroup);
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: setGroup.group_id,
    };
  });

export const deleteSetGroupCommand = (bundle, exerciseId, groupId) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    findSetGroup(exercise, groupId);
    ensure(exercise.set_groups.length > 1, "At least 1 set group");
    exercise.set_groups = exercise.set_groups.filter(
      (group) => group.group_id !== groupId
    );
    exercise.set_groups.forEach((group) => {
      if (group.backoff_config?.depends_on_set_group_id === groupId) {
        group.backoff_config.depends_on_set_group_id = "";
      }
    });
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
    };
  });

export const setSetGroupVariableCommand = (
  bundle,
  exerciseId,
  groupId,
  variableParameter
) =>
  applyProgramMutation(bundle, (working) => {
    ensure(
      VARIABLE_PARAMETERS.includes(variableParameter),
      "variable-parameter must be one of reps, weight, rpe"
    );
    const { exercise } = findExercise(working, exerciseId);
    const { group } = findSetGroup(exercise, groupId);
    ensure(
      variableParameter === "rpe" ||
        (!group.backoff_config?.enabled && !group.fatigue_drop_config?.enabled),
      "Backoff and fatigue drop require variable_parameter to remain 'rpe'"
    );
    ensure(
      variableParameter !== "weight" || !group.mix_weight_config?.enabled,
      "mixed weight model must be disabled before weight can be the variable parameter"
    );
    const duration = getExerciseDuration(working, exercise);
    group.variable_parameter = variableParameter;
    clearVariableParameterField(group, variableParameter, duration);
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
    };
  });

export const setWeekValueCommand = (
  bundle,
  exerciseId,
  groupId,
  weekIndex,
  field,
  value,
  subField
) =>
  applyProgramMutation(bundle, (working) => {
    ensure(WEEK_VALUE_FIELDS.includes(field), `Unsupported week value field: ${field}`);
    const { exercise } = findExercise(working, exerciseId);
    const { group } = findSetGroup(exercise, groupId);
    ensureWeekValueFieldIsEditable(group, field);
    const duration = ensureWeekIndexInExercise(working, exercise, weekIndex);
    const fallback = getDefaultForWeekValueField(field, subField);

    switch (field) {
      case "weekly_notes":
        setWeeklyArrayValue(
          group,
          "weekly_notes",
          weekIndex,
          safeSanitizeNote(String(value ?? "")),
          duration,
          fallback
        );
        break;
      case "mix_weight_config": {
        ensure(
          MIX_WEIGHT_SUB_FIELDS.includes(subField),
          `mix_weight_config requires --sub-field ${MIX_WEIGHT_SUB_FIELDS.join("|")}`
        );
        group.mix_weight_config = group.mix_weight_config || {
          enabled: false,
          weight_unit: "kg",
        };
        if (subField === "weight_unit") {
          group.mix_weight_config.weight_unit = String(value).trim().toLowerCase();
        } else {
          setWeeklyArrayValue(
            group.mix_weight_config,
            subField,
            weekIndex,
            parseNumericFlagValue(value, subField),
            duration,
            fallback
          );
        }
        break;
      }
      case "backoff_%":
        ensure(group.backoff_config?.enabled, "backoff_config must be enabled");
        setWeeklyArrayValue(
          group.backoff_config,
          "weekly_percentage",
          weekIndex,
          parseNumericFlagValue(value, field),
          duration,
          fallback
        );
        break;
      case "backoff_rpe":
        ensure(group.backoff_config?.enabled, "backoff_config must be enabled");
        setWeeklyArrayValue(
          group.backoff_config,
          "weekly_rpe",
          weekIndex,
          parseNumericFlagValue(value, field),
          duration,
          fallback
        );
        break;
      case "drop_%":
        ensure(
          group.fatigue_drop_config?.enabled,
          "fatigue_drop_config must be enabled"
        );
        setWeeklyArrayValue(
          group.fatigue_drop_config,
          "weekly_percentage",
          weekIndex,
          parseNumericFlagValue(value, field),
          duration,
          fallback
        );
        break;
      case "drop_rpe":
        ensure(
          group.fatigue_drop_config?.enabled,
          "fatigue_drop_config must be enabled"
        );
        setWeeklyArrayValue(
          group.fatigue_drop_config,
          "weekly_rpe",
          weekIndex,
          parseNumericFlagValue(value, field),
          duration,
          fallback
        );
        break;
      case "rpe_cap":
        ensure(
          group.fatigue_drop_config?.enabled,
          "fatigue_drop_config must be enabled"
        );
        setWeeklyArrayValue(
          group.fatigue_drop_config,
          "rpe_cap",
          weekIndex,
          parseNumericFlagValue(value, field),
          duration,
          fallback
        );
        break;
      default:
        setWeeklyArrayValue(
          group,
          field,
          weekIndex,
          parseNumericFlagValue(value, field),
          duration,
          fallback
        );
        break;
    }

    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
      week_index: weekIndex,
      field,
      ...(subField ? { sub_field: subField } : {}),
    };
  });

export const toggleBackoffCommand = (bundle, exerciseId, groupId, enabled) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    const { group, groupIndex } = findSetGroup(exercise, groupId);
    ensure(groupIndex > 0 || !enabled, "Cannot enable backoff on the first set group");
    if (enabled) {
      requireRpeVariableParameter(group, "Backoff");
    }
    const duration = getExerciseDuration(working, exercise);
    group.backoff_config = group.backoff_config || {
      enabled: false,
      depends_on_set_group_id: "",
      type: "",
    };
    group.backoff_config.enabled = Boolean(enabled);

    if (enabled) {
      group.backoff_config.type = group.backoff_config.type || "%_of_weight";
      ensureArrayValueSlot(group.backoff_config, "weekly_percentage", duration, 80);
      ensureArrayValueSlot(group.backoff_config, "weekly_rpe", duration, 8);
    } else {
      group.backoff_config.type = "";
    }

    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
    };
  });

export const setBackoffSourceCommand = (
  bundle,
  exerciseId,
  groupId,
  sourceGroupId
) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    const { group, groupIndex } = findSetGroup(exercise, groupId);
    const { groupIndex: sourceIndex } = findSetGroup(exercise, sourceGroupId);
    ensure(sourceIndex < groupIndex, "Backoff source must be an earlier set group");
    group.backoff_config = group.backoff_config || {
      enabled: false,
      depends_on_set_group_id: "",
      type: "",
    };
    group.backoff_config.depends_on_set_group_id = sourceGroupId;
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
      source_group_id: sourceGroupId,
    };
  });

export const setBackoffTypeCommand = (bundle, exerciseId, groupId, type) =>
  applyProgramMutation(bundle, (working) => {
    const normalizedType = normalizeDynamicWeightType(type);
    ensure(
      DYNAMIC_WEIGHT_TYPES.includes(normalizedType),
      "backoff type must be one of %_of_weight, target_rpe, or empty"
    );
    const { exercise } = findExercise(working, exerciseId);
    const { group } = findSetGroup(exercise, groupId);
    const duration = getExerciseDuration(working, exercise);
    group.backoff_config = group.backoff_config || {
      enabled: false,
      depends_on_set_group_id: "",
      type: "",
    };
    group.backoff_config.type = normalizedType;
    if (normalizedType === "%_of_weight") {
      ensureArrayValueSlot(group.backoff_config, "weekly_percentage", duration, 80);
    }
    if (normalizedType === "target_rpe") {
      ensureArrayValueSlot(group.backoff_config, "weekly_rpe", duration, 8);
    }
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
      type: normalizedType,
    };
  });

export const toggleFatigueDropCommand = (bundle, exerciseId, groupId, enabled) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    const { group } = findSetGroup(exercise, groupId);
    if (enabled) {
      requireRpeVariableParameter(group, "Fatigue drop");
    }
    const duration = getExerciseDuration(working, exercise);
    group.fatigue_drop_config = group.fatigue_drop_config || {
      enabled: false,
      type: "",
      rpe_cap: [],
    };
    group.fatigue_drop_config.enabled = Boolean(enabled);

    if (enabled) {
      group.fatigue_drop_config.type =
        group.fatigue_drop_config.type || "%_of_weight";
      ensureArrayValueSlot(group.fatigue_drop_config, "weekly_percentage", duration, 80);
      ensureArrayValueSlot(group.fatigue_drop_config, "weekly_rpe", duration, 8);
      ensureArrayValueSlot(group.fatigue_drop_config, "rpe_cap", duration, 8.5);
    } else {
      group.fatigue_drop_config.type = "";
    }

    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
    };
  });

export const setFatigueDropTypeCommand = (bundle, exerciseId, groupId, type) =>
  applyProgramMutation(bundle, (working) => {
    const normalizedType = normalizeDynamicWeightType(type);
    ensure(
      DYNAMIC_WEIGHT_TYPES.includes(normalizedType),
      "fatigue drop type must be one of %_of_weight, target_rpe, or empty"
    );
    const { exercise } = findExercise(working, exerciseId);
    const { group } = findSetGroup(exercise, groupId);
    const duration = getExerciseDuration(working, exercise);
    group.fatigue_drop_config = group.fatigue_drop_config || {
      enabled: false,
      type: "",
      rpe_cap: [],
    };
    group.fatigue_drop_config.type = normalizedType;
    if (normalizedType === "%_of_weight") {
      ensureArrayValueSlot(
        group.fatigue_drop_config,
        "weekly_percentage",
        duration,
        80
      );
    }
    if (normalizedType === "target_rpe") {
      ensureArrayValueSlot(group.fatigue_drop_config, "weekly_rpe", duration, 8);
    }
    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
      type: normalizedType,
    };
  });

export const setWeightModelCommand = (bundle, exerciseId, groupId, mode) =>
  applyProgramMutation(bundle, (working) => {
    ensure(
      WEIGHT_MODEL_MODES.includes(mode),
      "weight model must be one of percentage or mixed"
    );
    const { exercise } = findExercise(working, exerciseId);
    const { group } = findSetGroup(exercise, groupId);
    ensure(
      mode !== "mixed" || group.variable_parameter !== "weight",
      "mixed weight model cannot be enabled when weight is the variable parameter"
    );
    const duration = getExerciseDuration(working, exercise);
    group.mix_weight_config = group.mix_weight_config || {
      enabled: false,
      weight_unit: "kg",
    };
    group.mix_weight_config.enabled = mode === "mixed";

    if (mode === "mixed") {
      ensureArrayValueSlot(
        group.mix_weight_config,
        "weekly_weight_percentage",
        duration,
        70
      );
      ensureArrayValueSlot(
        group.mix_weight_config,
        "weekly_weight_absolute",
        duration,
        0
      );
    } else {
      ensureArrayValueSlot(group, "weekly_weight_percentage", duration, 65);
    }

    exercise.updated_at = nowIso();

    return {
      exercise_id: exerciseId,
      group_id: groupId,
      mode,
    };
  });

export const deleteExerciseCommand = (bundle, blockId, dayId, exerciseId) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);
    day.exercises = day.exercises.filter((currentExerciseId) => currentExerciseId !== exerciseId);
    working.exercises = working.exercises.filter((exercise) => exercise._id !== exerciseId);

    return {
      exercise_id: exerciseId,
    };
  });

export const moveExerciseCommand = (
  bundle,
  blockId,
  dayId,
  exerciseId,
  direction
) =>
  applyProgramMutation(bundle, (working) => {
    ensure(["up", "down", "left", "right"].includes(direction), "Direction must be one of up, down, left, right");
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);

    if (direction === "up" || direction === "down") {
      const currentIndex = day.exercises.indexOf(exerciseId);
      ensure(currentIndex !== -1, `Exercise ${exerciseId} not found in day ${dayId}`);

      if (direction === "up" && currentIndex > 0) {
        day.exercises.splice(currentIndex, 1);
        day.exercises.splice(currentIndex - 1, 0, exerciseId);
      }

      if (direction === "down" && currentIndex < day.exercises.length - 1) {
        day.exercises.splice(currentIndex, 1);
        day.exercises.splice(currentIndex + 1, 0, exerciseId);
      }
    } else {
      const orderedDayIds = getOrderedTrainingDayIds(block);
      const currentDayIndex = orderedDayIds.indexOf(dayId);
      ensure(currentDayIndex !== -1, `Training day ${dayId} not found in weekly schedule`);
      const targetDayIndex =
        direction === "left" ? currentDayIndex - 1 : currentDayIndex + 1;
      ensure(targetDayIndex >= 0 && targetDayIndex < orderedDayIds.length, `Cannot move exercise ${direction} from day ${dayId}`);

      const targetDayId = orderedDayIds[targetDayIndex];
      const { day: targetDay } = findDay(block, targetDayId);
      day.exercises = day.exercises.filter((currentExerciseId) => currentExerciseId !== exerciseId);
      targetDay.exercises.push(exerciseId);
    }

    return {
      exercise_id: exerciseId,
      direction,
    };
  });

export const copyExerciseCommand = (bundle, blockId, dayId, exerciseId) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);
    const { exercise } = findExercise(working, exerciseId);
    const clonedExercise = deepClone(exercise);
    clonedExercise._id = generateId("ex");
    clonedExercise.exercise_name = safeSanitizeName(
      `${exercise.exercise_name} _copy ${randomSuffix()}`
    );
    clonedExercise.created_at = nowIso();
    clonedExercise.updated_at = clonedExercise.created_at;
    remapSetGroups(clonedExercise, {});
    working.exercises.push(clonedExercise);
    day.exercises.push(clonedExercise._id);

    return {
      exercise_id: clonedExercise._id,
      source_exercise_id: exerciseId,
    };
  });

export const copyBlockCommand = (bundle, blockId, addRandomSuffix = true) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const clonedBlock = deepClone(block);
    const blockExercises = working.exercises.filter((exercise) => exercise.block_id === blockId);
    const clonedExercises = deepClone(blockExercises);
    const idMap = {
      days: {},
      exercises: {},
      setGroups: {},
    };

    clonedBlock._id = generateId("block");
    clonedBlock.name = safeSanitizeName(
      addRandomSuffix ? `${block.name} _copy ${randomSuffix()}` : block.name
    );

    clonedBlock.training_days.forEach((day) => {
      const oldDayId = day._id;
      const newDayId = generateId("day");
      idMap.days[oldDayId] = newDayId;
      day._id = newDayId;
    });

    clonedBlock.weekly_schedule = clonedBlock.weekly_schedule.map((dayId) =>
      dayId && idMap.days[dayId] ? idMap.days[dayId] : dayId
    );

    clonedExercises.forEach((exercise) => {
      const oldExerciseId = exercise._id;
      const newExerciseId = generateId("ex");
      idMap.exercises[oldExerciseId] = newExerciseId;
      exercise._id = newExerciseId;
      exercise.block_id = clonedBlock._id;
      exercise.created_at = nowIso();
      exercise.updated_at = exercise.created_at;
      remapSetGroups(exercise, idMap.setGroups);
    });

    clonedBlock.training_days.forEach((day) => {
      day.exercises = day.exercises.map((exerciseRef) => idMap.exercises[exerciseRef] || exerciseRef);
    });

    working.program.blocks.push(clonedBlock);
    working.exercises.push(...clonedExercises);

    return {
      block_id: clonedBlock._id,
      source_block_id: blockId,
    };
  });

export const copyDayCommand = (bundle, blockId, dayId) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const emptyIndex = block.weekly_schedule.lastIndexOf("");
    ensure(emptyIndex !== -1, "Cannot copy day because weekly_schedule is full");

    const { day } = findDay(block, dayId);
    const sourceExercises = day.exercises
      .map((exerciseId) => findExercise(working, exerciseId).exercise)
      .map((exercise) => deepClone(exercise));
    const clonedDay = deepClone(day);
    const idMap = {
      exercises: {},
      setGroups: {},
    };

    clonedDay._id = generateId("day");
    clonedDay.name = safeSanitizeName(`${day.name} _copy ${randomSuffix()}`);

    sourceExercises.forEach((exercise) => {
      const oldExerciseId = exercise._id;
      const newExerciseId = generateId("ex");
      idMap.exercises[oldExerciseId] = newExerciseId;
      exercise._id = newExerciseId;
      exercise.created_at = nowIso();
      exercise.updated_at = exercise.created_at;
      remapSetGroups(exercise, idMap.setGroups);
    });

    clonedDay.exercises = day.exercises.map((exerciseId) => idMap.exercises[exerciseId] || exerciseId);
    block.training_days.push(clonedDay);
    block.weekly_schedule[emptyIndex] = clonedDay._id;
    working.exercises.push(...sourceExercises);

    return {
      day_id: clonedDay._id,
      source_day_id: dayId,
      position: emptyIndex,
    };
  });

export const addCustomLiftCommand = (bundle, liftName) =>
  applyProgramMutation(bundle, (working) => {
    working.program.custom_anchored_lifts =
      working.program.custom_anchored_lifts || [];
    const validation = validateCustomLiftName(
      liftName,
      working.program.custom_anchored_lifts
    );
    ensure(validation.valid, validation.error);
    working.program.custom_anchored_lifts.push(validation.trimmed);

    return {
      lift_name: validation.trimmed,
    };
  });

export const renameCustomLiftCommand = (bundle, oldLiftName, newLiftName) =>
  applyProgramMutation(bundle, (working) => {
    const customLifts = working.program.custom_anchored_lifts || [];
    const index = customLifts.findIndex(
      (lift) => lift.toLowerCase() === oldLiftName.toLowerCase()
    );
    ensure(index !== -1, `Custom lift ${oldLiftName} not found`);

    const validation = validateCustomLiftName(
      newLiftName,
      customLifts.filter((_, candidateIndex) => candidateIndex !== index)
    );
    ensure(validation.valid, validation.error);
    const exactOldName = customLifts[index];
    customLifts[index] = validation.trimmed;

    working.exercises.forEach((exercise) => {
      const liftType = exercise.one_rm_anchor?.lift_type;
      if (
        typeof liftType === "string" &&
        liftType.toLowerCase() === exactOldName.toLowerCase()
      ) {
        exercise.one_rm_anchor.lift_type = validation.trimmed;
      }
    });

    return {
      old_lift_name: exactOldName,
      new_lift_name: validation.trimmed,
    };
  });

export const deleteCustomLiftCommand = (bundle, liftName) =>
  applyProgramMutation(bundle, (working) => {
    const customLifts = working.program.custom_anchored_lifts || [];
    const index = customLifts.findIndex(
      (lift) => lift.toLowerCase() === liftName.toLowerCase()
    );
    ensure(index !== -1, `Custom lift ${liftName} not found`);
    const removedName = customLifts[index];
    customLifts.splice(index, 1);

    working.exercises.forEach((exercise) => {
      const liftType = exercise.one_rm_anchor?.lift_type;
      if (
        typeof liftType === "string" &&
        liftType.toLowerCase() === removedName.toLowerCase()
      ) {
        exercise.one_rm_anchor = { enabled: false };
      }
    });

    return {
      lift_name: removedName,
    };
  });

export const createProcessFromProgramCommand = (bundle, options = {}) =>
  createProcessBundleFromProgram(bundle, options);

export const updateProcessCommand = (bundle, patch = {}) =>
  applyProcessMutation(bundle, (working) => {
    const safePatch = pickAllowedPatch(
      patch,
      PROCESS_UPDATE_FIELDS,
      "Process update payload"
    );
    working.process = {
      ...working.process,
      ...safePatch,
      name: safePatch.name ? safeSanitizeName(safePatch.name) : working.process.name,
    };

    return {
      process_id: working.process._id,
    };
  });

export const updateProcessConfigCommand = (bundle, configPatch = {}) =>
  applyProcessMutation(bundle, (working) => {
    const safeConfigPatch = pickAllowedPatch(
      configPatch,
      PROCESS_CONFIG_FIELDS,
      "Process config payload"
    );
    working.process.config = {
      ...working.process.config,
      ...safeConfigPatch,
    };

    return {
      process_id: working.process._id,
    };
  });

export const updateProcessOneRmCommand = (bundle, oneRmProfile) =>
  applyProcessMutation(bundle, (working) => {
    validateOneRmProfileKeys(working.program, oneRmProfile, "One RM profile payload");
    working.process.one_rm_profile = mergeOneRmProfile(
      working.process.one_rm_profile,
      oneRmProfile
    );

    return {
      process_id: working.process._id,
    };
  });

export const setProcessOneRmCommand = (
  bundle,
  liftName,
  value,
  blockIndex = null
) =>
  applyProcessMutation(bundle, (working) => {
    const liftKey = normalizeLiftKey(liftName);
    ensure(liftKey.length > 0, "lift must be a non-empty string");
    ensure(
      getAllowedOneRmLiftKeys(working.program).has(liftKey),
      `Unsupported lift key: ${liftName}. Add the custom lift to the program first.`
    );
    const numericValue = parseNumericFlagValue(value, "one RM value");

    if (blockIndex === null || blockIndex === undefined) {
      working.process.one_rm_profile = mergeOneRmProfile(
        working.process.one_rm_profile,
        { [liftKey]: numericValue }
      );
    } else {
      ensure(
        Number.isInteger(blockIndex) && blockIndex >= 0,
        "block-index must be a non-negative integer"
      );
      ensure(
        blockIndex < working.program.blocks.length,
        `block-index ${blockIndex} is outside the program block count (${working.program.blocks.length})`
      );
      if (!Array.isArray(working.process.one_rm_profile.blockly_one_rm)) {
        working.process.one_rm_profile.blockly_one_rm = [];
      }
      while (working.process.one_rm_profile.blockly_one_rm.length <= blockIndex) {
        working.process.one_rm_profile.blockly_one_rm.push({});
      }
      working.process.one_rm_profile.blockly_one_rm[blockIndex] = {
        ...working.process.one_rm_profile.blockly_one_rm[blockIndex],
        [liftKey]: numericValue,
      };
    }

    return {
      process_id: working.process._id,
      lift: liftKey,
      ...(blockIndex === null || blockIndex === undefined
        ? {}
        : { block_index: blockIndex }),
    };
  });

export const updateProcessConfigAndOneRmCommand = (
  bundle,
  oneRmProfile,
  configPatch
) =>
  applyProcessMutation(bundle, (working) => {
    const safeConfigPatch = pickAllowedPatch(
      configPatch,
      PROCESS_CONFIG_FIELDS,
      "Process config payload"
    );
    validateOneRmProfileKeys(working.program, oneRmProfile, "One RM profile payload");
    working.process.one_rm_profile = mergeOneRmProfile(
      working.process.one_rm_profile,
      oneRmProfile
    );
    working.process.config = {
      weight_unit: "kg",
      week_start_day: 0,
      weight_rounding: 2.5,
      ...working.process.config,
      ...safeConfigPatch,
    };

    return {
      process_id: working.process._id,
    };
  });

export const updateProcessSetCommand = (
  bundle,
  exerciseId,
  weekIndex,
  setIndex,
  setData
) =>
  applyProcessMutation(bundle, (working) => {
    ensure(Number.isInteger(weekIndex) && weekIndex >= 0, "week-index must be a non-negative integer");
    ensure(Number.isInteger(setIndex) && setIndex >= 0, "set-index must be a non-negative integer");
    const safeSetData = pickAllowedPatch(
      setData,
      PROCESS_SET_FIELDS,
      "Process set payload"
    );
    const recording =
      working.process.exercise_recordings.find(
        (candidate) => candidate.exercise_id === exerciseId
      ) || null;
    ensure(recording, `Exercise recording for ${exerciseId} not found`);
    const blockDuration = getExerciseBlockDuration(working, exerciseId);
    ensure(
      weekIndex < blockDuration,
      `week-index ${weekIndex} is outside the exercise block duration (${blockDuration})`
    );
    const prescribedSetCount = getPrescribedSetCount(working, exerciseId, weekIndex);
    ensure(
      setIndex < prescribedSetCount,
      `set-index ${setIndex} is outside the prescribed set count (${prescribedSetCount}) for week ${weekIndex}`
    );

    while (recording.weekly.length <= weekIndex) {
      recording.weekly.push({ sets: [], note: "" });
    }

    while (recording.weekly[weekIndex].sets.length <= setIndex) {
      recording.weekly[weekIndex].sets.push({
        weight: -1,
        reps: -1,
        rpe: -1,
        completed: false,
      });
    }

    recording.weekly[weekIndex].sets[setIndex] = {
      ...(recording.weekly[weekIndex].sets[setIndex] || {
        weight: -1,
        reps: -1,
        rpe: -1,
        completed: false,
      }),
      ...safeSetData,
    };

    return {
      process_id: working.process._id,
      exercise_id: exerciseId,
      week_index: weekIndex,
      set_index: setIndex,
    };
  });

export const updateProcessNoteCommand = (bundle, exerciseId, weekIndex, note) =>
  applyProcessMutation(bundle, (working) => {
    ensure(Number.isInteger(weekIndex) && weekIndex >= 0, "week-index must be a non-negative integer");
    const recording =
      working.process.exercise_recordings.find(
        (candidate) => candidate.exercise_id === exerciseId
      ) || null;
    ensure(recording, `Exercise recording for ${exerciseId} not found`);
    const blockDuration = getExerciseBlockDuration(working, exerciseId);
    ensure(
      weekIndex < blockDuration,
      `week-index ${weekIndex} is outside the exercise block duration (${blockDuration})`
    );

    while (recording.weekly.length <= weekIndex) {
      recording.weekly.push({ sets: [], note: "" });
    }

    recording.weekly[weekIndex].note = safeSanitizeNote(note);

    return {
      process_id: working.process._id,
      exercise_id: exerciseId,
      week_index: weekIndex,
    };
  });

export const getExerciseRecordingCommand = (bundle, exerciseId) => {
  const normalized = normalizeProcessBundle(bundle).bundle;
  const recording =
    normalized.process.exercise_recordings.find(
      (candidate) => candidate.exercise_id === exerciseId
    ) || null;
  ensure(recording, `Exercise recording for ${exerciseId} not found`);

  return {
    recording,
    warnings: [],
  };
};
