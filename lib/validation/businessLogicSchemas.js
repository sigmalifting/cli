import { secureObjectValidation } from "./shared.js";

const ValidationResult = {
  success: (data, warnings = []) => ({ success: true, data, warnings }),
  error: (errors, details = null) => ({ success: false, errors, details }),
};

const STANDARD_ANCHOR_LIFTS = ["squat", "bench", "deadlift"];
const VARIABLE_PARAMETER_FIELDS = {
  reps: "weekly_reps",
  rpe: "weekly_rpe",
  weight: "weekly_weight_percentage",
};

const normalizeAnchorLiftName = (liftName) =>
  typeof liftName === "string" ? liftName.trim().toLowerCase() : "";

const buildAllowedAnchorLiftNames = (program) => [
  ...STANDARD_ANCHOR_LIFTS,
  ...((program?.custom_anchored_lifts || []).filter(
    (liftName) => typeof liftName === "string" && liftName.trim() !== ""
  )),
];

const validateBlockReferences = (program, exercises) => {
  const warnings = [];

  if (!program?.blocks) {
    return ValidationResult.error([
      {
        path: "program.blocks",
        message: "Program blocks are required",
        code: "missing_blocks",
      },
    ]);
  }

  if (!exercises || !Array.isArray(exercises)) {
    return ValidationResult.error([
      {
        path: "exercises",
        message: "Exercises array is required",
        code: "missing_exercises",
      },
    ]);
  }

  const cleanedProgram = JSON.parse(JSON.stringify(program));
  const cleanedExercises = JSON.parse(JSON.stringify(exercises));
  const seenBlockIds = new Set();
  const originalCount = cleanedProgram.blocks.length;

  cleanedProgram.blocks = cleanedProgram.blocks.filter((block, blockIndex) => {
    if (seenBlockIds.has(block._id)) {
      warnings.push({
        path: `blocks[${blockIndex}]`,
        message: `Removed duplicate block ID: ${block._id} (only keeping first appearance)`,
        code: "duplicate_block_id",
      });
      return false;
    }

    seenBlockIds.add(block._id);
    return true;
  });

  if (cleanedProgram.blocks.length !== originalCount) {
    warnings.push({
      path: "blocks",
      message: `Cleaned ${
        originalCount - cleanedProgram.blocks.length
      } duplicate block references`,
      code: "cleaned_block_references",
    });
  }

  cleanedProgram.blocks.forEach((block) => {
    if (!block.training_days) {
      return;
    }

    block.training_days.forEach((day) => {
      if (!day.exercises || !Array.isArray(day.exercises)) {
        return;
      }

      day.exercises.forEach((exerciseId) => {
        const exercise = cleanedExercises.find((ex) => ex._id === exerciseId);
        if (exercise && exercise.block_id !== block._id) {
          exercise.block_id = block._id;
          warnings.push({
            path: `exercises[${cleanedExercises.findIndex((ex) => ex._id === exerciseId)}]`,
            message: `Enforced block_id assignment: ${exerciseId} -> ${block._id}`,
            code: "enforced_block_id_assignment",
          });
        }
      });
    });
  });

  return ValidationResult.success(
    { program: cleanedProgram, exercises: cleanedExercises },
    warnings
  );
};

const validateExerciseReferences = (program, exercises) => {
  const warnings = [];

  if (!program?.blocks || !Array.isArray(exercises)) {
    return ValidationResult.error([
      {
        path: "root",
        message: "Invalid program or exercises data",
        code: "invalid_structure",
      },
    ]);
  }

  const exerciseIds = new Set(exercises.map((ex) => ex._id));
  const cleanedProgram = JSON.parse(JSON.stringify(program));
  const cleanedExercises = JSON.parse(JSON.stringify(exercises));
  const referencedExerciseIds = new Set();

  cleanedProgram.blocks.forEach((block, blockIndex) => {
    if (!block.training_days) {
      return;
    }

    block.training_days.forEach((day, dayIndex) => {
      if (!day.exercises || !Array.isArray(day.exercises)) {
        return;
      }

      const originalCount = day.exercises.length;
      day.exercises = day.exercises.filter((exerciseId) => {
        if (!exerciseIds.has(exerciseId)) {
          warnings.push({
            path: `blocks[${blockIndex}].training_days[${dayIndex}].exercises`,
            message: `Removed invalid exercise reference: ${exerciseId}`,
            code: "invalid_exercise_reference",
          });
          return false;
        }

        if (referencedExerciseIds.has(exerciseId)) {
          warnings.push({
            path: `blocks[${blockIndex}].training_days[${dayIndex}].exercises`,
            message: `Removed duplicate exercise reference: ${exerciseId} (only keeping first appearance)`,
            code: "duplicate_exercise_reference",
          });
          return false;
        }

        referencedExerciseIds.add(exerciseId);
        return true;
      });

      if (day.exercises.length !== originalCount) {
        warnings.push({
          path: `blocks[${blockIndex}].training_days[${dayIndex}]`,
          message: `Cleaned ${
            originalCount - day.exercises.length
          } invalid/duplicate exercise references`,
          code: "cleaned_references",
        });
      }
    });
  });

  const originalExerciseCount = cleanedExercises.length;
  const filteredExercises = cleanedExercises.filter((exercise) => {
    if (!referencedExerciseIds.has(exercise._id)) {
      warnings.push({
        path: "exercises",
        message: `Removed orphaned exercise definition: ${exercise._id} (${
          exercise.exercise_name || "unnamed"
        })`,
        code: "orphaned_exercise_definition",
      });
      return false;
    }
    return true;
  });

  if (filteredExercises.length !== originalExerciseCount) {
    warnings.push({
      path: "exercises",
      message: `Cleaned ${
        originalExerciseCount - filteredExercises.length
      } orphaned exercise definitions`,
      code: "cleaned_exercise_definitions",
    });
  }

  return ValidationResult.success(
    { program: cleanedProgram, exercises: filteredExercises },
    warnings
  );
};

const validateTrainingDayReferences = (program) => {
  const warnings = [];

  if (!program?.blocks) {
    return ValidationResult.error([
      {
        path: "program.blocks",
        message: "Program blocks are required",
        code: "missing_blocks",
      },
    ]);
  }

  const cleanedProgram = JSON.parse(JSON.stringify(program));

  cleanedProgram.blocks.forEach((block, blockIndex) => {
    if (!block.training_days || !block.weekly_schedule) {
      return;
    }

    const seenTrainingDayIds = new Set();
    const originalTrainingDayCount = block.training_days.length;

    block.training_days = block.training_days.filter((day, dayIndex) => {
      if (seenTrainingDayIds.has(day._id)) {
        warnings.push({
          path: `blocks[${blockIndex}].training_days[${dayIndex}]`,
          message: `Removed duplicate training day ID: ${day._id} (only keeping first appearance)`,
          code: "duplicate_training_day_id",
        });
        return false;
      }

      seenTrainingDayIds.add(day._id);
      return true;
    });

    if (block.training_days.length !== originalTrainingDayCount) {
      warnings.push({
        path: `blocks[${blockIndex}].training_days`,
        message: `Cleaned ${
          originalTrainingDayCount - block.training_days.length
        } duplicate training day references`,
        code: "cleaned_training_day_references",
      });
    }

    const trainingDayIds = new Set(block.training_days.map((day) => day._id));

    block.weekly_schedule = block.weekly_schedule.map((dayId, scheduleIndex) => {
      if (dayId && dayId !== "" && !trainingDayIds.has(dayId)) {
        warnings.push({
          path: `blocks[${blockIndex}].weekly_schedule[${scheduleIndex}]`,
          message: `Removed invalid training day reference: ${dayId}`,
          code: "invalid_training_day_reference",
        });
        return "";
      }
      return dayId;
    });

    const scheduledIds = new Set(
      block.weekly_schedule.filter((id) => id && id !== "")
    );
    block.training_days.forEach((day, dayIndex) => {
      if (!scheduledIds.has(day._id)) {
        warnings.push({
          path: `blocks[${blockIndex}].training_days[${dayIndex}]`,
          message: `Training day ${day._id} (${day.name}) exists but is not scheduled`,
          code: "orphaned_training_day",
        });
      }
    });
  });

  return ValidationResult.success(cleanedProgram, warnings);
};

const validateSetGroupReferences = (exercises) => {
  const warnings = [];
  const cleanedExercises = JSON.parse(JSON.stringify(exercises));

  cleanedExercises.forEach((exercise, exerciseIndex) => {
    if (!exercise.set_groups || !Array.isArray(exercise.set_groups)) {
      return;
    }

    const seenSetGroupIds = new Set();
    const originalCount = exercise.set_groups.length;

    exercise.set_groups = exercise.set_groups.filter((setGroup, setGroupIndex) => {
      if (seenSetGroupIds.has(setGroup.group_id)) {
        warnings.push({
          path: `exercises[${exerciseIndex}].set_groups[${setGroupIndex}]`,
          message: `Removed duplicate set group ID: ${setGroup.group_id} (only keeping first appearance)`,
          code: "duplicate_set_group_id",
        });
        return false;
      }

      seenSetGroupIds.add(setGroup.group_id);
      return true;
    });

    if (exercise.set_groups.length !== originalCount) {
      warnings.push({
        path: `exercises[${exerciseIndex}].set_groups`,
        message: `Cleaned ${
          originalCount - exercise.set_groups.length
        } duplicate set group references`,
        code: "cleaned_set_group_references",
      });
    }
  });

  return ValidationResult.success(cleanedExercises, warnings);
};

const validateSetGroupDependencies = (exercises) => {
  const warnings = [];
  const cleanedExercises = JSON.parse(JSON.stringify(exercises));

  cleanedExercises.forEach((exercise, exerciseIndex) => {
    if (!exercise.set_groups || !Array.isArray(exercise.set_groups)) {
      return;
    }

    exercise.set_groups.forEach((setGroup, setGroupIndex) => {
      if (
        setGroup.backoff_config?.enabled &&
        setGroup.backoff_config.depends_on_set_group_id
      ) {
        const dependsOnId = setGroup.backoff_config.depends_on_set_group_id;
        const dependencyIndex = exercise.set_groups.findIndex(
          (sg) => sg.group_id === dependsOnId
        );

        if (dependencyIndex === -1) {
          setGroup.backoff_config.depends_on_set_group_id = "";
          warnings.push({
            path: `exercises[${exerciseIndex}].set_groups[${setGroupIndex}].backoff_config`,
            message: `Invalid backoff dependency ${dependsOnId} - dependency not found`,
            code: "invalid_backoff_dependency",
          });
        } else if (dependencyIndex >= setGroupIndex) {
          setGroup.backoff_config.depends_on_set_group_id = "";
          warnings.push({
            path: `exercises[${exerciseIndex}].set_groups[${setGroupIndex}].backoff_config`,
            message: `Invalid backoff dependency ${dependsOnId} - dependency must come before current set group`,
            code: "invalid_dependency_order",
          });
        }
      }

      if (
        setGroupIndex === 0 &&
        setGroup.backoff_config?.enabled &&
        setGroup.backoff_config.depends_on_set_group_id
      ) {
        setGroup.backoff_config.enabled = false;
        setGroup.backoff_config.depends_on_set_group_id = "";
        warnings.push({
          path: `exercises[${exerciseIndex}].set_groups[0].backoff_config`,
          message:
            "First set group cannot be a backoff set - disabled backoff configuration",
          code: "first_set_cannot_be_backoff",
        });
      }
    });
  });

  return ValidationResult.success(cleanedExercises, warnings);
};

const validateOneRmAnchorReferences = (program, exercises) => {
  const errors = [];
  const allowedLiftNames = buildAllowedAnchorLiftNames(program);
  const allowedLiftKeys = new Set(allowedLiftNames.map(normalizeAnchorLiftName));
  const allowedLiftList = allowedLiftNames.join(", ");

  if (!Array.isArray(exercises)) {
    return ValidationResult.error([
      {
        path: "exercises",
        message: "Exercises array is required",
        code: "missing_exercises",
      },
    ]);
  }

  exercises.forEach((exercise, exerciseIndex) => {
    const anchor = exercise.one_rm_anchor;
    if (!anchor?.enabled) {
      return;
    }

    const liftType =
      typeof anchor.lift_type === "string" ? anchor.lift_type.trim() : "";
    if (liftType === "") {
      errors.push({
        path: `exercises[${exerciseIndex}].one_rm_anchor.lift_type`,
        message:
          "Enabled 1RM anchor requires a lift_type selected from squat, bench, deadlift, or program.custom_anchored_lifts",
        code: "missing_one_rm_anchor_lift_type",
      });
      return;
    }

    if (!allowedLiftKeys.has(normalizeAnchorLiftName(liftType))) {
      errors.push({
        path: `exercises[${exerciseIndex}].one_rm_anchor.lift_type`,
        message: `1RM anchor lift_type '${liftType}' must be one of: ${allowedLiftList}`,
        code: "invalid_one_rm_anchor_lift_type",
      });
    }
  });

  if (errors.length > 0) {
    return ValidationResult.error(errors);
  }

  return ValidationResult.success(exercises);
};

const validateVariableParameterFields = (exercises) => {
  const warnings = [];
  const cleanedExercises = JSON.parse(JSON.stringify(exercises));

  if (!Array.isArray(cleanedExercises)) {
    return ValidationResult.error([
      {
        path: "exercises",
        message: "Exercises array is required",
        code: "missing_exercises",
      },
    ]);
  }

  cleanedExercises.forEach((exercise, exerciseIndex) => {
    if (!Array.isArray(exercise.set_groups)) {
      return;
    }

    exercise.set_groups.forEach((setGroup, setGroupIndex) => {
      const field = VARIABLE_PARAMETER_FIELDS[setGroup.variable_parameter];
      const values = field ? setGroup[field] : null;

      if (!Array.isArray(values)) {
        return;
      }

      if (values.some((value) => value !== -1)) {
        setGroup[field] = values.map(() => -1);
        warnings.push({
          path: `exercises[${exerciseIndex}].set_groups[${setGroupIndex}].${field}`,
          message: `${field} is recorded in workout data when variable_parameter is '${setGroup.variable_parameter}' and was reset to -1`,
          code: "cleared_variable_parameter_values",
        });
      }
    });
  });

  return ValidationResult.success(cleanedExercises, warnings);
};

const validateVariableParameterSemantics = (exercises) => {
  const errors = [];
  const dynamicWeightTypes = new Set(["%_of_weight", "target_rpe"]);

  if (!Array.isArray(exercises)) {
    return ValidationResult.error([
      {
        path: "exercises",
        message: "Exercises array is required",
        code: "missing_exercises",
      },
    ]);
  }

  exercises.forEach((exercise, exerciseIndex) => {
    if (!Array.isArray(exercise.set_groups)) {
      return;
    }

    exercise.set_groups.forEach((setGroup, setGroupIndex) => {
      const pathPrefix = `exercises[${exerciseIndex}].set_groups[${setGroupIndex}]`;
      const variableParameter = setGroup.variable_parameter;

      if (setGroup.backoff_config?.enabled) {
        if (variableParameter !== "rpe") {
          errors.push({
            path: `${pathPrefix}.backoff_config`,
            message:
              "Backoff sets require variable_parameter to be 'rpe' so weight can be derived while RPE is recorded",
            code: "invalid_variable_parameter_backoff",
          });
        }

        if (!setGroup.backoff_config.depends_on_set_group_id) {
          errors.push({
            path: `${pathPrefix}.backoff_config.depends_on_set_group_id`,
            message: "Enabled backoff_config requires a source set group",
            code: "missing_backoff_dependency",
          });
        }

        if (!dynamicWeightTypes.has(setGroup.backoff_config.type)) {
          errors.push({
            path: `${pathPrefix}.backoff_config.type`,
            message:
              "Enabled backoff_config.type must be '%_of_weight' or 'target_rpe'",
            code: "missing_backoff_type",
          });
        }
      }

      if (setGroup.fatigue_drop_config?.enabled) {
        if (variableParameter !== "rpe") {
          errors.push({
            path: `${pathPrefix}.fatigue_drop_config`,
            message:
              "Fatigue drops require variable_parameter to be 'rpe' so RPE can trigger the drop",
            code: "invalid_variable_parameter_fatigue_drop",
          });
        }

        if (!dynamicWeightTypes.has(setGroup.fatigue_drop_config.type)) {
          errors.push({
            path: `${pathPrefix}.fatigue_drop_config.type`,
            message:
              "Enabled fatigue_drop_config.type must be '%_of_weight' or 'target_rpe'",
            code: "missing_fatigue_drop_type",
          });
        }
      }

      if (
        variableParameter === "weight" &&
        setGroup.mix_weight_config?.enabled
      ) {
        errors.push({
          path: `${pathPrefix}.mix_weight_config`,
          message:
            "mix_weight_config prescribes weight and cannot be enabled when weight is the variable parameter",
          code: "invalid_variable_parameter_mix_weight",
        });
      }
    });
  });

  if (errors.length > 0) {
    return ValidationResult.error(errors);
  }

  return ValidationResult.success(exercises);
};

const validateProcessProgramConsistency = (process, program, exercises) => {
  const warnings = [];
  const validationErrors = [];

  if (!process || !program || !Array.isArray(exercises)) {
    return ValidationResult.error([
      {
        path: "root",
        message: "Process, program, and exercises are all required",
        code: "missing_required_data",
      },
    ]);
  }

  if (process.program_id !== program._id) {
    validationErrors.push({
      path: "process.program_id",
      message: `Process program_id (${process.program_id}) doesn't match program._id (${program._id})`,
      code: "program_id_mismatch",
    });
  }

  const cleanedProcess = JSON.parse(JSON.stringify(process));
  const exerciseIds = new Set(exercises.map((ex) => ex._id));

  if (
    !cleanedProcess.exercise_recordings ||
    !Array.isArray(cleanedProcess.exercise_recordings)
  ) {
    cleanedProcess.exercise_recordings = [];
  }

  const originalCount = cleanedProcess.exercise_recordings.length;
  cleanedProcess.exercise_recordings =
    cleanedProcess.exercise_recordings.filter((recording) => {
      if (!exerciseIds.has(recording.exercise_id)) {
        warnings.push({
          path: "process.exercise_recordings",
          message: `Removed recording for non-existent exercise: ${recording.exercise_id}`,
          code: "invalid_exercise_recording",
        });
        return false;
      }
      return true;
    });

  if (cleanedProcess.exercise_recordings.length !== originalCount) {
    warnings.push({
      path: "process.exercise_recordings",
      message: `Cleaned ${
        originalCount - cleanedProcess.exercise_recordings.length
      } invalid exercise recordings`,
      code: "cleaned_exercise_recordings",
    });
  }

  const referencedExerciseIds = new Set();
  program.blocks?.forEach((block) => {
    block.training_days?.forEach((day) => {
      day.exercises?.forEach((exerciseId) => {
        referencedExerciseIds.add(exerciseId);
      });
    });
  });

  const existingRecordingIds = new Set(
    cleanedProcess.exercise_recordings.map((recording) => recording.exercise_id)
  );
  const missingRecordingIds = [...referencedExerciseIds].filter(
    (entryId) => !existingRecordingIds.has(entryId)
  );

  if (missingRecordingIds.length > 0) {
    const blockDurations = {};
    program.blocks?.forEach((block) => {
      if (block.duration) {
        blockDurations[block._id] = block.duration;
      }
    });

    missingRecordingIds.forEach((exerciseId) => {
      const exercise = exercises.find((ex) => ex._id === exerciseId);
      const blockDuration = blockDurations[exercise.block_id];
      const weekly = Array.from({ length: blockDuration }, () => ({
        sets: [],
        note: "",
      }));

      cleanedProcess.exercise_recordings.push({
        exercise_id: exerciseId,
        weekly,
      });
    });

    warnings.push({
      path: "process.exercise_recordings",
      message: `Created ${
        missingRecordingIds.length
      } missing exercise recordings for: ${missingRecordingIds.join(", ")}`,
      code: "created_missing_exercise_recordings",
    });
  }

  if (cleanedProcess.one_rm_profile?.blockly_one_rm && program.blocks) {
    const expectedBlockCount = program.blocks.length;
    const actualBlockCount =
      cleanedProcess.one_rm_profile.blockly_one_rm.length;

    if (actualBlockCount !== expectedBlockCount) {
      if (actualBlockCount > expectedBlockCount) {
        cleanedProcess.one_rm_profile.blockly_one_rm =
          cleanedProcess.one_rm_profile.blockly_one_rm.slice(
            0,
            expectedBlockCount
          );
      } else {
        const customLiftsForBlock = {};
        if (
          program.custom_anchored_lifts &&
          Array.isArray(program.custom_anchored_lifts)
        ) {
          program.custom_anchored_lifts.forEach((liftName) => {
            customLiftsForBlock[liftName.toLowerCase()] = -1;
          });
        }

        while (
          cleanedProcess.one_rm_profile.blockly_one_rm.length <
          expectedBlockCount
        ) {
          cleanedProcess.one_rm_profile.blockly_one_rm.push({
            squat: -1,
            bench: -1,
            deadlift: -1,
            ...customLiftsForBlock,
          });
        }
      }

      warnings.push({
        path: "process.one_rm_profile.blockly_one_rm",
        message: `Adjusted blockly_one_rm array from ${actualBlockCount} to ${expectedBlockCount} blocks`,
        code: "adjusted_blockly_one_rm",
      });
    }
  }

  if (validationErrors.length > 0) {
    return ValidationResult.error(validationErrors);
  }

  return ValidationResult.success(cleanedProcess, warnings);
};

const validateConfigArrayLengths = (
  config,
  configName,
  arrayNames,
  expectedDuration,
  exerciseIndex,
  setGroupIndex,
  warnings
) => {
  if (!config) {
    return;
  }

  arrayNames.forEach((arrayName) => {
    const array = config[arrayName];
    if (!array || !Array.isArray(array)) {
      return;
    }

    const currentLength = array.length;
    if (currentLength === expectedDuration) {
      return;
    }

    if (currentLength > expectedDuration) {
      config[arrayName] = array.slice(0, expectedDuration);
    } else {
      while (config[arrayName].length < expectedDuration) {
        config[arrayName].push(-1);
      }
    }

    warnings.push({
      path: `exercises[${exerciseIndex}].set_groups[${setGroupIndex}].${configName}.${arrayName}`,
      message: `Adjusted ${configName}.${arrayName} length from ${currentLength} to ${expectedDuration}`,
      code: "adjusted_config_array_length",
    });
  });
};

const validateWeeklyArrayLengths = (exercises, program) => {
  const warnings = [];
  const cleanedExercises = JSON.parse(JSON.stringify(exercises));
  const blockDurations = {};

  if (program?.blocks) {
    program.blocks.forEach((block) => {
      blockDurations[block._id] = block.duration || 4;
    });
  }

  cleanedExercises.forEach((exercise, exerciseIndex) => {
    const expectedDuration = blockDurations[exercise.block_id] || 4;

    if (!exercise.set_groups || !Array.isArray(exercise.set_groups)) {
      return;
    }

    exercise.set_groups.forEach((setGroup, setGroupIndex) => {
      const weeklyArrays = [
        "weekly_notes",
        "weekly_num_sets",
        "weekly_reps",
        "weekly_rpe",
        "weekly_weight_percentage",
      ];

      weeklyArrays.forEach((arrayName) => {
        if (setGroup[arrayName] && Array.isArray(setGroup[arrayName])) {
          const currentLength = setGroup[arrayName].length;

          if (currentLength !== expectedDuration) {
            if (currentLength > expectedDuration) {
              setGroup[arrayName] = setGroup[arrayName].slice(
                0,
                expectedDuration
              );
            } else {
              const fillValue = arrayName === "weekly_notes" ? "" : -1;
              while (setGroup[arrayName].length < expectedDuration) {
                setGroup[arrayName].push(fillValue);
              }
            }

            warnings.push({
              path: `exercises[${exerciseIndex}].set_groups[${setGroupIndex}].${arrayName}`,
              message: `Adjusted ${arrayName} length from ${currentLength} to ${expectedDuration}`,
              code: "adjusted_weekly_array_length",
            });
          }
        }
      });

      validateConfigArrayLengths(
        setGroup.mix_weight_config,
        "mix_weight_config",
        ["weekly_weight_percentage", "weekly_weight_absolute"],
        expectedDuration,
        exerciseIndex,
        setGroupIndex,
        warnings
      );

      validateConfigArrayLengths(
        setGroup.backoff_config,
        "backoff_config",
        ["weekly_rpe", "weekly_percentage"],
        expectedDuration,
        exerciseIndex,
        setGroupIndex,
        warnings
      );

      validateConfigArrayLengths(
        setGroup.fatigue_drop_config,
        "fatigue_drop_config",
        ["rpe_cap", "weekly_rpe", "weekly_percentage"],
        expectedDuration,
        exerciseIndex,
        setGroupIndex,
        warnings
      );
    });
  });

  return ValidationResult.success(cleanedExercises, warnings);
};

const validateFinalConsistency = (cleanedData, type) => {
  const errors = [];
  const referencedExerciseIds = new Set();

  cleanedData.program?.blocks?.forEach((block) => {
    block.training_days?.forEach((day) => {
      day.exercises?.forEach((exerciseId) => {
        referencedExerciseIds.add(exerciseId);
      });
    });
  });

  const providedExerciseIds = new Set(
    cleanedData.exercises?.map((ex) => ex._id) || []
  );
  const recordedExerciseIds = new Set();

  if (type === "process" && cleanedData.process?.exercise_recordings) {
    cleanedData.process.exercise_recordings.forEach((recording) => {
      recordedExerciseIds.add(recording.exercise_id);
    });
  }

  if (referencedExerciseIds.size !== providedExerciseIds.size) {
    errors.push({
      path: "root",
      message: `Count mismatch: Program references ${referencedExerciseIds.size} exercises but ${providedExerciseIds.size} definitions provided`,
      code: "exercise_count_mismatch",
    });
  }

  const missingDefinitions = [...referencedExerciseIds].filter(
    (entryId) => !providedExerciseIds.has(entryId)
  );
  if (missingDefinitions.length > 0) {
    errors.push({
      path: "exercises",
      message: `Missing exercise definitions for: ${missingDefinitions.join(", ")}`,
      code: "missing_exercise_definitions",
    });
  }

  const orphanedDefinitions = [...providedExerciseIds].filter(
    (entryId) => !referencedExerciseIds.has(entryId)
  );
  if (orphanedDefinitions.length > 0) {
    errors.push({
      path: "exercises",
      message: `Orphaned exercise definitions: ${orphanedDefinitions.join(", ")}`,
      code: "orphaned_exercise_definitions",
    });
  }

  if (type === "process") {
    const extraRecordings = [...recordedExerciseIds].filter(
      (entryId) => !referencedExerciseIds.has(entryId)
    );
    if (extraRecordings.length > 0) {
      errors.push({
        path: "process.exercise_recordings",
        message: `Process has recordings for exercises not in program: ${extraRecordings.join(", ")}`,
        code: "extra_exercise_recordings",
      });
    }
  }

  if (type === "process") {
    const missingRecordings = [...referencedExerciseIds].filter(
      (entryId) => !recordedExerciseIds.has(entryId)
    );
    if (missingRecordings.length > 0) {
      errors.push({
        path: "process.exercise_recordings",
        message: `Process missing recordings for program exercises: ${missingRecordings.join(", ")}`,
        code: "missing_exercise_recordings",
      });
    }
  }

  if (errors.length > 0) {
    return ValidationResult.error(errors);
  }

  return ValidationResult.success(cleanedData);
};

const validateImportBusinessLogic = (importData, type = "program") => {
  const securityCheck = secureObjectValidation(importData);
  if (!securityCheck.valid) {
    return ValidationResult.error([
      {
        path: "root",
        message: securityCheck.error,
        code: "security_violation",
      },
    ]);
  }

  const allWarnings = [];
  const allErrors = [];
  let cleanedData = JSON.parse(JSON.stringify(importData));

  try {
    if (cleanedData.program && cleanedData.exercises) {
      const blockResult = validateBlockReferences(
        cleanedData.program,
        cleanedData.exercises
      );
      if (!blockResult.success) {
        allErrors.push(...blockResult.errors);
      } else {
        cleanedData.program = blockResult.data.program;
        cleanedData.exercises = blockResult.data.exercises;
        allWarnings.push(...blockResult.warnings);
      }
    }

    if (cleanedData.program) {
      const trainingDayResult = validateTrainingDayReferences(cleanedData.program);
      if (!trainingDayResult.success) {
        allErrors.push(...trainingDayResult.errors);
      } else {
        cleanedData.program = trainingDayResult.data;
        allWarnings.push(...trainingDayResult.warnings);
      }
    }

    if (cleanedData.program && cleanedData.exercises) {
      const exerciseRefResult = validateExerciseReferences(
        cleanedData.program,
        cleanedData.exercises
      );
      if (!exerciseRefResult.success) {
        allErrors.push(...exerciseRefResult.errors);
      } else {
        cleanedData.program = exerciseRefResult.data.program;
        cleanedData.exercises = exerciseRefResult.data.exercises;
        allWarnings.push(...exerciseRefResult.warnings);
      }
    }

    if (cleanedData.exercises) {
      const setGroupRefResult = validateSetGroupReferences(cleanedData.exercises);
      if (!setGroupRefResult.success) {
        allErrors.push(...setGroupRefResult.errors);
      } else {
        cleanedData.exercises = setGroupRefResult.data;
        allWarnings.push(...setGroupRefResult.warnings);
      }
    }

    if (cleanedData.exercises) {
      const setGroupResult = validateSetGroupDependencies(cleanedData.exercises);
      if (!setGroupResult.success) {
        allErrors.push(...setGroupResult.errors);
      } else {
        cleanedData.exercises = setGroupResult.data;
        allWarnings.push(...setGroupResult.warnings);
      }
    }

    if (cleanedData.exercises) {
      const variableParameterFieldResult = validateVariableParameterFields(
        cleanedData.exercises
      );
      if (!variableParameterFieldResult.success) {
        allErrors.push(...variableParameterFieldResult.errors);
      } else {
        cleanedData.exercises = variableParameterFieldResult.data;
        allWarnings.push(...variableParameterFieldResult.warnings);
      }
    }

    if (cleanedData.exercises) {
      const variableParameterResult = validateVariableParameterSemantics(
        cleanedData.exercises
      );
      if (!variableParameterResult.success) {
        allErrors.push(...variableParameterResult.errors);
      } else {
        cleanedData.exercises = variableParameterResult.data;
        allWarnings.push(...variableParameterResult.warnings);
      }
    }

    if (cleanedData.exercises && cleanedData.program) {
      const weeklyArrayResult = validateWeeklyArrayLengths(
        cleanedData.exercises,
        cleanedData.program
      );
      if (!weeklyArrayResult.success) {
        allErrors.push(...weeklyArrayResult.errors);
      } else {
        cleanedData.exercises = weeklyArrayResult.data;
        allWarnings.push(...weeklyArrayResult.warnings);
      }
    }

    if (cleanedData.exercises && cleanedData.program) {
      const anchorResult = validateOneRmAnchorReferences(
        cleanedData.program,
        cleanedData.exercises
      );
      if (!anchorResult.success) {
        allErrors.push(...anchorResult.errors);
      } else {
        cleanedData.exercises = anchorResult.data;
        allWarnings.push(...anchorResult.warnings);
      }
    }

    if (type === "process" && cleanedData.process) {
      const processResult = validateProcessProgramConsistency(
        cleanedData.process,
        cleanedData.program,
        cleanedData.exercises
      );
      if (!processResult.success) {
        allErrors.push(...processResult.errors);
      } else {
        cleanedData.process = processResult.data;
        allWarnings.push(...processResult.warnings);
      }
    }

    const finalConsistencyResult = validateFinalConsistency(cleanedData, type);
    if (!finalConsistencyResult.success) {
      allErrors.push(...finalConsistencyResult.errors);
    }

    if (allErrors.length > 0) {
      return ValidationResult.error(allErrors);
    }

    return ValidationResult.success(cleanedData, allWarnings);
  } catch (error) {
    return ValidationResult.error([
      {
        path: "root",
        message: `Validation failed: ${error.message}`,
        code: "validation_exception",
      },
    ]);
  }
};

export {
  ValidationResult,
  validateBlockReferences,
  validateTrainingDayReferences,
  validateExerciseReferences,
  validateSetGroupReferences,
  validateSetGroupDependencies,
  validateOneRmAnchorReferences,
  validateVariableParameterFields,
  validateVariableParameterSemantics,
  validateProcessProgramConsistency,
  validateWeeklyArrayLengths,
  validateFinalConsistency,
  validateImportBusinessLogic,
};
