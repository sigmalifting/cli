import {
  normalizeProcessBundle,
  normalizeProgramBundle,
} from "./bundleOps.js";

const WEEKDAY_LABELS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

export const INSPECT_VIEWS = [
  "program-overview",
  "process-overview",
  "exercise-detail",
  "recording-detail",
  "one-rm-profile",
];
export const RENDER_VIEWS = [...INSPECT_VIEWS];

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

const detectBundleKind = (bundle) => {
  if (bundle?.process && bundle?.program && Array.isArray(bundle?.exercises)) {
    return "process-import";
  }

  if (bundle?.program && Array.isArray(bundle?.exercises) && !bundle?.process) {
    return "program-import";
  }

  throw createCliError(
    "Unsupported bundle shape. Expected a program-import or process-import payload."
  );
};

const normalizeBundleForInspect = (bundle) => {
  const bundleKind = detectBundleKind(bundle);
  return {
    bundleKind,
    bundle:
      bundleKind === "process-import"
        ? normalizeProcessBundle(bundle).bundle
        : normalizeProgramBundle(bundle).bundle,
  };
};

const summarizeSetStatus = (recording) => {
  const weeklyEntries = recording?.weekly || [];
  let totalSets = 0;
  let completedSets = 0;
  let notedWeeks = 0;

  weeklyEntries.forEach((week) => {
    const sets = week?.sets || [];
    totalSets += sets.length;
    completedSets += sets.filter((set) => set.completed).length;
    if (week?.note) {
      notedWeeks += 1;
    }
  });

  return {
    week_count: weeklyEntries.length,
    total_sets: totalSets,
    completed_sets: completedSets,
    noted_weeks: notedWeeks,
  };
};

const sortObjectEntries = (value) =>
  Object.fromEntries(
    Object.entries(value || {}).sort(([left], [right]) => left.localeCompare(right))
  );

const buildBundleContext = (normalizedBundle, bundleKind) => {
  const program = normalizedBundle.program;
  const process = bundleKind === "process-import" ? normalizedBundle.process : null;
  const exercises = normalizedBundle.exercises || [];

  const exerciseById = new Map(exercises.map((exercise) => [exercise._id, exercise]));
  const recordingByExerciseId = new Map(
    (process?.exercise_recordings || []).map((recording) => [
      recording.exercise_id,
      recording,
    ])
  );
  const dayById = new Map();
  const blockById = new Map();
  const locationByExerciseId = new Map();

  program.blocks.forEach((block, blockIndex) => {
    blockById.set(block._id, {
      ...block,
      block_index: blockIndex,
    });

    block.training_days.forEach((day) => {
      dayById.set(day._id, {
        ...day,
        block_id: block._id,
        block_index: blockIndex,
      });

      day.exercises.forEach((exerciseId, orderInDay) => {
        locationByExerciseId.set(exerciseId, {
          block_id: block._id,
          block_name: block.name,
          block_index: blockIndex,
          day_id: day._id,
          day_name: day.name,
          order_in_day: orderInDay,
        });
      });
    });
  });

  return {
    bundleKind,
    bundle: normalizedBundle,
    program,
    process,
    exercises,
    exerciseById,
    recordingByExerciseId,
    dayById,
    blockById,
    locationByExerciseId,
  };
};

const buildBlockSchedule = (block, dayById) =>
  (block.weekly_schedule || []).map((dayId, position) => ({
    position,
    weekday: WEEKDAY_LABELS[position] || `Day ${position}`,
    day_id: dayId || "",
    occupied: Boolean(dayId),
    day_name: dayId ? dayById.get(dayId)?.name || "" : "",
  }));

const findExerciseContext = (context, exerciseId) => {
  const exercise = context.exerciseById.get(exerciseId) || null;
  ensure(exercise, `Exercise ${exerciseId} not found`);

  const location = context.locationByExerciseId.get(exerciseId) || null;
  ensure(location, `Exercise ${exerciseId} does not belong to any training day`);

  const block = context.blockById.get(location.block_id) || null;
  ensure(block, `Block ${location.block_id} not found for exercise ${exerciseId}`);

  const day = context.dayById.get(location.day_id) || null;
  ensure(day, `Day ${location.day_id} not found for exercise ${exerciseId}`);

  return {
    exercise,
    location,
    block,
    day,
  };
};

const buildProgramOverviewPage = (context) => ({
  view: "program-overview",
  title: context.program.name,
  description: context.program.description || "",
  bundle_kind: context.bundleKind,
  custom_lifts: [...(context.program.custom_anchored_lifts || [])],
  summary: {
    block_count: context.program.blocks.length,
    training_day_count: context.program.blocks.reduce(
      (count, block) => count + (block.training_days?.length || 0),
      0
    ),
    exercise_count: context.exercises.length,
  },
  blocks: context.program.blocks.map((block, blockIndex) => ({
    block_id: block._id,
    block_index: blockIndex,
    name: block.name,
    duration: block.duration,
    schedule: buildBlockSchedule(block, context.dayById),
    days: (block.training_days || []).map((day) => ({
      day_id: day._id,
      name: day.name,
      weekly_positions: buildBlockSchedule(block, context.dayById)
        .filter((entry) => entry.day_id === day._id)
        .map((entry) => entry.position),
      exercise_count: day.exercises.length,
      exercises: day.exercises.map((exerciseId) => {
        const exercise = context.exerciseById.get(exerciseId);
        return {
          exercise_id: exerciseId,
          name: exercise?.exercise_name || "Unknown Exercise",
          variable_parameters: (exercise?.set_groups || []).map(
            (setGroup) => setGroup.variable_parameter
          ),
        };
      }),
    })),
  })),
});

const buildProcessOverviewPage = (context) => {
  ensure(
    context.process,
    "process-overview requires a process-import bundle"
  );

  const recordingSummaries = [...context.recordingByExerciseId.values()].map((recording) =>
    summarizeSetStatus(recording)
  );

  return {
    view: "process-overview",
    title: context.process.name,
    bundle_kind: context.bundleKind,
    process: {
      process_id: context.process._id,
      program_id: context.process.program_id,
      program_name: context.process.program_name,
      start_date: context.process.start_date,
    },
    config: { ...(context.process.config || {}) },
    one_rm_profile_summary: {
      enable_blockly_one_rm: Boolean(
        context.process.one_rm_profile?.enable_blockly_one_rm
      ),
      block_count: (context.process.one_rm_profile?.blockly_one_rm || []).length,
    },
    summary: {
      block_count: context.program.blocks.length,
      exercise_count: context.exercises.length,
      recording_count: context.process.exercise_recordings.length,
      total_sets: recordingSummaries.reduce(
        (count, entry) => count + entry.total_sets,
        0
      ),
      completed_sets: recordingSummaries.reduce(
        (count, entry) => count + entry.completed_sets,
        0
      ),
      noted_weeks: recordingSummaries.reduce(
        (count, entry) => count + entry.noted_weeks,
        0
      ),
    },
    blocks: context.program.blocks.map((block, blockIndex) => {
      const blockExerciseIds = new Set(
        context.exercises
          .filter((exercise) => exercise.block_id === block._id)
          .map((exercise) => exercise._id)
      );
      const blockRecordings = [...blockExerciseIds]
        .map((exerciseId) => context.recordingByExerciseId.get(exerciseId))
        .filter(Boolean);
      const blockStats = blockRecordings.map((recording) =>
        summarizeSetStatus(recording)
      );

      return {
        block_id: block._id,
        block_index: blockIndex,
        name: block.name,
        duration: block.duration,
        day_count: block.training_days.length,
        exercise_count: blockExerciseIds.size,
        schedule: buildBlockSchedule(block, context.dayById),
        recording_stats: {
          recording_count: blockRecordings.length,
          total_sets: blockStats.reduce(
            (count, entry) => count + entry.total_sets,
            0
          ),
          completed_sets: blockStats.reduce(
            (count, entry) => count + entry.completed_sets,
            0
          ),
          noted_weeks: blockStats.reduce(
            (count, entry) => count + entry.noted_weeks,
            0
          ),
        },
      };
    }),
  };
};

const buildExerciseDetailPage = (context, exerciseId) => {
  const { exercise, location } = findExerciseContext(context, exerciseId);

  return {
    view: "exercise-detail",
    bundle_kind: context.bundleKind,
    exercise: {
      exercise_id: exercise._id,
      name: exercise.exercise_name,
      block_id: location.block_id,
      block_name: location.block_name,
      block_index: location.block_index,
      day_id: location.day_id,
      day_name: location.day_name,
      order_in_day: location.order_in_day,
      one_rm_anchor: exercise.one_rm_anchor || { enabled: false },
      deload_config: exercise.deload_config || { enabled: false },
      set_group_count: (exercise.set_groups || []).length,
    },
    set_groups: (exercise.set_groups || []).map((setGroup, index) => ({
      index,
      group_id: setGroup.group_id,
      variable_parameter: setGroup.variable_parameter,
      weekly_num_sets: [...(setGroup.weekly_num_sets || [])],
      weekly_reps: [...(setGroup.weekly_reps || [])],
      weekly_rpe: [...(setGroup.weekly_rpe || [])],
      weekly_weight_percentage: [...(setGroup.weekly_weight_percentage || [])],
      weekly_notes: [...(setGroup.weekly_notes || [])],
      mix_weight_config: setGroup.mix_weight_config
        ? { ...setGroup.mix_weight_config }
        : null,
      backoff_config: setGroup.backoff_config
        ? { ...setGroup.backoff_config }
        : null,
      fatigue_drop_config: setGroup.fatigue_drop_config
        ? { ...setGroup.fatigue_drop_config }
        : null,
    })),
  };
};

const buildRecordingDetailPage = (context, exerciseId) => {
  ensure(
    context.process,
    "recording-detail requires a process-import bundle"
  );

  const { exercise, location } = findExerciseContext(context, exerciseId);
  const recording = context.recordingByExerciseId.get(exerciseId) || null;
  ensure(recording, `Recording for exercise ${exerciseId} not found`);

  return {
    view: "recording-detail",
    bundle_kind: context.bundleKind,
    exercise: {
      exercise_id: exercise._id,
      name: exercise.exercise_name,
      block_name: location.block_name,
      day_name: location.day_name,
    },
    summary: summarizeSetStatus(recording),
    weeks: (recording.weekly || []).map((week, weekIndex) => ({
      week_index: weekIndex,
      note: week.note || "",
      set_count: (week.sets || []).length,
      completed_set_count: (week.sets || []).filter((set) => set.completed)
        .length,
      sets: (week.sets || []).map((set, setIndex) => ({
        set_index: setIndex,
        weight: set.weight,
        reps: set.reps,
        rpe: set.rpe,
        completed: Boolean(set.completed),
      })),
    })),
  };
};

const buildOneRmProfilePage = (context) => {
  ensure(
    context.process,
    "one-rm-profile requires a process-import bundle"
  );

  const oneRmProfile = context.process.one_rm_profile || {};
  const { blockly_one_rm = [], enable_blockly_one_rm, ...globalValues } =
    oneRmProfile;

  return {
    view: "one-rm-profile",
    bundle_kind: context.bundleKind,
    process: {
      process_id: context.process._id,
      name: context.process.name,
      start_date: context.process.start_date,
    },
    config: { ...(context.process.config || {}) },
    enable_blockly_one_rm: Boolean(enable_blockly_one_rm),
    global_values: sortObjectEntries(globalValues),
    block_values: blockly_one_rm.map((values, blockIndex) => ({
      block_index: blockIndex,
      block_name: context.program.blocks[blockIndex]?.name || `Block ${blockIndex + 1}`,
      values: sortObjectEntries(values),
    })),
  };
};

export const buildInspectPage = (bundle, view, options = {}) => {
  const { bundleKind, bundle: normalizedBundle } = normalizeBundleForInspect(bundle);
  const context = buildBundleContext(normalizedBundle, bundleKind);

  switch (view) {
    case "program-overview":
      return {
        bundleKind,
        page: buildProgramOverviewPage(context),
      };
    case "process-overview":
      return {
        bundleKind,
        page: buildProcessOverviewPage(context),
      };
    case "exercise-detail":
      ensure(options.exerciseId, "exercise-detail requires --exercise-id");
      return {
        bundleKind,
        page: buildExerciseDetailPage(context, options.exerciseId),
      };
    case "recording-detail":
      ensure(options.exerciseId, "recording-detail requires --exercise-id");
      return {
        bundleKind,
        page: buildRecordingDetailPage(context, options.exerciseId),
      };
    case "one-rm-profile":
      return {
        bundleKind,
        page: buildOneRmProfilePage(context),
      };
    default:
      throw createCliError(`Unknown inspect view: ${view}`);
  }
};

const divider = (char = "-", width = 72) => char.repeat(width);

const formatKeyValue = (label, value) => `${label}: ${value}`;

const formatValue = (value) => {
  if (Array.isArray(value)) {
    return `[${value.join(", ")}]`;
  }

  if (value === null || value === undefined || value === "") {
    return "-";
  }

  if (typeof value === "object") {
    return JSON.stringify(value);
  }

  return String(value);
};

const formatScheduleLine = (schedule) =>
  schedule
    .map((entry) =>
      `${entry.weekday}:${entry.occupied ? entry.day_name || entry.day_id : "-"}`
    )
    .join(" | ");

const formatVariableParameters = (value) =>
  value && value.length > 0 ? value.join("/") : "-";

const renderProgramOverviewPage = (page) => {
  const lines = [
    divider("="),
    `Program Overview: ${page.title}`,
    divider("="),
    formatKeyValue("Description", page.description || "-"),
    formatKeyValue(
      "Summary",
      `${page.summary.block_count} blocks | ${page.summary.training_day_count} days | ${page.summary.exercise_count} exercises`
    ),
    formatKeyValue(
      "Custom Lifts",
      page.custom_lifts.length > 0 ? page.custom_lifts.join(", ") : "none"
    ),
  ];

  page.blocks.forEach((block) => {
    lines.push("");
    lines.push(
      `[Block ${block.block_index + 1}] ${block.name} (${block.duration} weeks)`
    );
    lines.push(formatKeyValue("Schedule", formatScheduleLine(block.schedule)));

    block.days.forEach((day) => {
      lines.push(
        `  Day: ${day.name} [${day.weekly_positions
          .map((position) => WEEKDAY_LABELS[position])
          .join(", ")}]`
      );
      day.exercises.forEach((exercise, index) => {
        lines.push(
          `    ${index + 1}. ${exercise.name} (${formatVariableParameters(
            exercise.variable_parameters
          )})`
        );
      });
    });
  });

  return lines.join("\n");
};

const renderProcessOverviewPage = (page) => {
  const lines = [
    divider("="),
    `Process Overview: ${page.title}`,
    divider("="),
    formatKeyValue("Program", page.process.program_name),
    formatKeyValue("Start Date", page.process.start_date),
    formatKeyValue(
      "Config",
      `unit ${page.config.weight_unit || "-"} | rounding ${
        page.config.weight_rounding ?? "-"
      } | week start ${page.config.week_start_day ?? "-"}`
    ),
    formatKeyValue(
      "Summary",
      `${page.summary.block_count} blocks | ${page.summary.exercise_count} exercises | ${page.summary.recording_count} recordings | ${page.summary.completed_sets} completed sets`
    ),
    formatKeyValue(
      "Blockly 1RM",
      page.one_rm_profile_summary.enable_blockly_one_rm
        ? `enabled (${page.one_rm_profile_summary.block_count} blocks)`
        : `disabled (${page.one_rm_profile_summary.block_count} blocks)`
    ),
  ];

  page.blocks.forEach((block) => {
    lines.push("");
    lines.push(
      `[Block ${block.block_index + 1}] ${block.name} (${block.duration} weeks)`
    );
    lines.push(
      formatKeyValue(
        "Coverage",
        `${block.day_count} days | ${block.exercise_count} exercises | ${block.recording_stats.recording_count} recordings | ${block.recording_stats.completed_sets} completed sets`
      )
    );
    lines.push(formatKeyValue("Schedule", formatScheduleLine(block.schedule)));
  });

  return lines.join("\n");
};

const renderExerciseDetailPage = (page) => {
  const lines = [
    divider("="),
    `Exercise Detail: ${page.exercise.name}`,
    divider("="),
    formatKeyValue(
      "Location",
      `Block ${page.exercise.block_index + 1} ${page.exercise.block_name} / ${page.exercise.day_name} / slot ${page.exercise.order_in_day + 1}`
    ),
    formatKeyValue(
      "Anchor",
      page.exercise.one_rm_anchor?.enabled
        ? `${page.exercise.one_rm_anchor.lift_type} x ${page.exercise.one_rm_anchor.ratio}`
        : "disabled"
    ),
    formatKeyValue(
      "Deload",
      page.exercise.deload_config?.enabled
        ? `enabled (${page.exercise.deload_config.percentage}%)`
        : "disabled"
    ),
    formatKeyValue("Set Groups", String(page.exercise.set_group_count)),
  ];

  page.set_groups.forEach((group) => {
    lines.push("");
    lines.push(`[Set Group ${group.index + 1}] ${group.variable_parameter}`);
    lines.push(`  Sets: ${formatValue(group.weekly_num_sets)}`);
    lines.push(`  Reps: ${formatValue(group.weekly_reps)}`);
    lines.push(`  RPE: ${formatValue(group.weekly_rpe)}`);
    lines.push(`  %: ${formatValue(group.weekly_weight_percentage)}`);
    lines.push(
      `  Notes: ${
        group.weekly_notes.some(Boolean)
          ? group.weekly_notes.map((note) => note || "-").join(" | ")
          : "none"
      }`
    );
    lines.push(
      `  Mix Weight: ${
        group.mix_weight_config?.enabled
          ? JSON.stringify(group.mix_weight_config)
          : "disabled"
      }`
    );
    lines.push(
      `  Backoff: ${
        group.backoff_config?.enabled
          ? JSON.stringify(group.backoff_config)
          : "disabled"
      }`
    );
    lines.push(
      `  Fatigue Drop: ${
        group.fatigue_drop_config?.enabled
          ? JSON.stringify(group.fatigue_drop_config)
          : "disabled"
      }`
    );
  });

  return lines.join("\n");
};

const renderRecordingDetailPage = (page) => {
  const lines = [
    divider("="),
    `Recording Detail: ${page.exercise.name}`,
    divider("="),
    formatKeyValue("Location", `${page.exercise.block_name} / ${page.exercise.day_name}`),
    formatKeyValue(
      "Summary",
      `${page.summary.week_count} weeks | ${page.summary.total_sets} sets | ${page.summary.completed_sets} completed | ${page.summary.noted_weeks} noted weeks`
    ),
  ];

  page.weeks.forEach((week) => {
    lines.push("");
    lines.push(`[Week ${week.week_index + 1}] note: ${week.note || "-"}`);
    if (week.sets.length === 0) {
      lines.push("  (no logged sets)");
      return;
    }

    week.sets.forEach((set) => {
      lines.push(
        `  ${set.set_index + 1}. ${set.weight} x ${set.reps} @ ${set.rpe} ${
          set.completed ? "done" : "open"
        }`
      );
    });
  });

  return lines.join("\n");
};

const renderOneRmProfilePage = (page) => {
  const lines = [
    divider("="),
    `1RM Profile: ${page.process.name}`,
    divider("="),
    formatKeyValue("Start Date", page.process.start_date),
    formatKeyValue(
      "Config",
      `unit ${page.config.weight_unit || "-"} | rounding ${
        page.config.weight_rounding ?? "-"
      } | week start ${page.config.week_start_day ?? "-"}`
    ),
    formatKeyValue(
      "Blockly 1RM",
      page.enable_blockly_one_rm ? "enabled" : "disabled"
    ),
    "",
    "Global Values",
  ];

  Object.entries(page.global_values).forEach(([key, value]) => {
    lines.push(`  ${key}: ${value}`);
  });

  page.block_values.forEach((block) => {
    lines.push("");
    lines.push(`[Block ${block.block_index + 1}] ${block.block_name}`);
    Object.entries(block.values).forEach(([key, value]) => {
      lines.push(`  ${key}: ${value}`);
    });
  });

  return lines.join("\n");
};

export const renderInspectPage = (page) => {
  switch (page.view) {
    case "program-overview":
      return renderProgramOverviewPage(page);
    case "process-overview":
      return renderProcessOverviewPage(page);
    case "exercise-detail":
      return renderExerciseDetailPage(page);
    case "recording-detail":
      return renderRecordingDetailPage(page);
    case "one-rm-profile":
      return renderOneRmProfilePage(page);
    default:
      throw createCliError(`Unknown render view: ${page.view}`);
  }
};
