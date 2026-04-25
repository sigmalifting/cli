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

  return {
    group_id: generateId("sg"),
    variable_parameter: variableParameter,
    weekly_notes: Array(weeks).fill(""),
    weekly_num_sets: Array(weeks).fill(3),
    weekly_reps: weeklyReps,
    weekly_rpe: weeklyRpe,
    weekly_weight_percentage: Array(weeks).fill(65),
    mix_weight_config: {
      enabled: false,
      weight_unit: "kg",
    },
    backoff_config: {
      enabled: false,
      depends_on_set_group_id: "",
      type: "",
    },
    fatigue_drop_config: {
      enabled: false,
      type: "",
      rpe_cap: [],
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

const normalizeBundleResult = (validation, kind) => {
  if (!validation.success) {
    throw createCliError(`Invalid ${kind} payload`, validation);
  }

  const businessValidation = validateImportBusinessLogic(validation.data, kind);
  if (!businessValidation.success) {
    throw createCliError(`Invalid ${kind} business logic`, businessValidation);
  }

  return {
    bundle: touchMetadata(businessValidation.data),
    warnings: businessValidation.warnings || [],
  };
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
    "program"
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
    "process"
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
      "Cross-reference cleanup and weekly-array normalization are handled by `normalize`.",
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

  (program.custom_anchored_lifts || []).forEach((liftName) => {
    customLifts[liftName.toLowerCase()] = -1;
  });

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
      ...config,
    },
    one_rm_profile: {
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
      ...oneRmProfile,
    },
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
    case "program-import":
      return validateProgramImport(data);
    case "process-import":
      return validateProcessImport(data);
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
    working.program = {
      ...working.program,
      ...patch,
      name: patch.name ? safeSanitizeName(patch.name) : working.program.name,
      description:
        patch.description !== undefined
          ? safeSanitizeDescription(patch.description)
          : working.program.description,
    };

    return {
      program_id: working.program._id,
    };
  });

export const addBlockCommand = (bundle, blockData = {}) =>
  applyProgramMutation(bundle, (working) => {
    const block = buildBlock({
      programId: working.program._id,
      name: blockData.name || "New Block",
      duration: blockData.duration || 4,
      weeklySchedule:
        blockData.weekly_schedule && blockData.weekly_schedule.length === 7
          ? blockData.weekly_schedule
          : [...EMPTY_WEEKLY_SCHEDULE],
      trainingDays: blockData.training_days || [],
    });

    working.program.blocks.push(block);

    return {
      block_id: block._id,
    };
  });

export const updateBlockCommand = (bundle, blockId, blockPatch = {}) =>
  applyProgramMutation(bundle, (working) => {
    const { block, blockIndex } = findBlock(working, blockId);
    working.program.blocks[blockIndex] = {
      ...block,
      ...blockPatch,
      name: blockPatch.name ? safeSanitizeName(blockPatch.name) : block.name,
      weekly_schedule:
        blockPatch.weekly_schedule && blockPatch.weekly_schedule.length === 7
          ? blockPatch.weekly_schedule
          : block.weekly_schedule,
    };

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
    ensure(Number.isInteger(position) && position >= 0 && position <= 6, "Day position must be between 0 and 6");
    const { block } = findBlock(working, blockId);
    ensure(!block.weekly_schedule[position], `Weekly schedule already has a day at position ${position}`);

    const day = buildTrainingDay({
      name: dayData.name || "New Day",
      exercises: dayData.exercises || [],
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
    block.training_days = block.training_days.filter((day) => day._id !== dayId);
    block.weekly_schedule = block.weekly_schedule.map((entry) =>
      entry === dayId ? "" : entry
    );

    return {
      day_id: dayId,
    };
  });

export const updateScheduleCommand = (bundle, blockId, weeklySchedule) =>
  applyProgramMutation(bundle, (working) => {
    ensure(Array.isArray(weeklySchedule) && weeklySchedule.length === 7, "Weekly schedule must be an array of 7 entries");
    const { block } = findBlock(working, blockId);
    block.weekly_schedule = weeklySchedule;

    return {
      block_id: blockId,
    };
  });

export const addExerciseCommand = (bundle, blockId, dayId, exerciseData = {}) =>
  applyProgramMutation(bundle, (working) => {
    const { block } = findBlock(working, blockId);
    const { day } = findDay(block, dayId);
    const exercise = buildExerciseDefinition({
      programId: working.program._id,
      blockId,
      exerciseName: exerciseData.exercise_name || exerciseData.name || "New Exercise",
      weeks: block.duration,
      variableParameter:
        exerciseData.variable_parameter ||
        exerciseData.variableParameter ||
        "weight",
      setGroups: exerciseData.set_groups,
      oneRmAnchor: exerciseData.one_rm_anchor,
      deloadConfig: exerciseData.deload_config,
    });

    working.exercises.push(exercise);
    day.exercises.push(exercise._id);

    return {
      exercise_id: exercise._id,
    };
  });

export const updateExerciseCommand = (bundle, exerciseId, patch = {}) =>
  applyProgramMutation(bundle, (working) => {
    const { exercise } = findExercise(working, exerciseId);
    Object.assign(exercise, patch);
    if (patch.exercise_name) {
      exercise.exercise_name = safeSanitizeName(patch.exercise_name);
    }
    exercise.updated_at = nowIso();

    const validation = validateExercise(exercise);
    if (!validation.success) {
      throw createCliError(`Updated exercise ${exerciseId} is invalid`, validation);
    }

    return {
      exercise_id: exerciseId,
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
    working.process = {
      ...working.process,
      ...patch,
      name: patch.name ? safeSanitizeName(patch.name) : working.process.name,
    };

    return {
      process_id: working.process._id,
    };
  });

export const updateProcessConfigCommand = (bundle, configPatch = {}) =>
  applyProcessMutation(bundle, (working) => {
    working.process.config = {
      ...working.process.config,
      ...configPatch,
    };

    return {
      process_id: working.process._id,
    };
  });

export const updateProcessOneRmCommand = (bundle, oneRmProfile) =>
  applyProcessMutation(bundle, (working) => {
    working.process.one_rm_profile = deepClone(oneRmProfile);

    return {
      process_id: working.process._id,
    };
  });

export const updateProcessConfigAndOneRmCommand = (
  bundle,
  oneRmProfile,
  configPatch
) =>
  applyProcessMutation(bundle, (working) => {
    working.process.one_rm_profile = deepClone(oneRmProfile);
    working.process.config = {
      weight_unit: "kg",
      week_start_day: 0,
      weight_rounding: 2.5,
      ...working.process.config,
      ...configPatch,
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
    const recording =
      working.process.exercise_recordings.find(
        (candidate) => candidate.exercise_id === exerciseId
      ) || null;
    ensure(recording, `Exercise recording for ${exerciseId} not found`);

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

    recording.weekly[weekIndex].sets[setIndex] = setData;

    return {
      process_id: working.process._id,
      exercise_id: exerciseId,
      week_index: weekIndex,
      set_index: setIndex,
    };
  });

export const updateProcessNoteCommand = (bundle, exerciseId, weekIndex, note) =>
  applyProcessMutation(bundle, (working) => {
    const recording =
      working.process.exercise_recordings.find(
        (candidate) => candidate.exercise_id === exerciseId
      ) || null;
    ensure(recording, `Exercise recording for ${exerciseId} not found`);

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
