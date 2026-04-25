import { z } from "zod";
import {
  idSchema,
  isoDateTimeSchema,
  weightSchema,
  weightDeltaSchema,
  repsSchema,
  rpeSchema,
  percentageSchema,
  anchorLiftSchema,
  anchorRatioSchema,
  limitedArray,
  optional,
  formatValidationErrors,
  safeSanitizeName,
  safeSanitizeNote,
} from "./shared.js";

const mixWeightConfigSchema = z.object({
  enabled: z.boolean(),
  weight_unit: z
    .string()
    .transform((str) => str.trim())
    .transform((str) => str.toLowerCase())
    .pipe(z.enum(["kg", "lbs"]))
    .optional(),
  weekly_weight_percentage: z.array(percentageSchema).optional(),
  weekly_weight_absolute: z.array(weightDeltaSchema).optional(),
});

const backoffConfigSchema = z.object({
  enabled: z.boolean(),
  depends_on_set_group_id: optional(idSchema),
  type: z
    .string()
    .transform((str) => str.trim())
    .transform((str) => str.toLowerCase())
    .pipe(z.enum(["%_of_weight", "target_rpe", ""]))
    .optional(),
  weekly_rpe: z.array(rpeSchema).optional(),
  weekly_percentage: z.array(percentageSchema).optional(),
});

const fatigueDropConfigSchema = z.object({
  enabled: z.boolean(),
  rpe_cap: z.array(rpeSchema).optional(),
  type: z
    .string()
    .transform((str) => str.trim())
    .transform((str) => str.toLowerCase())
    .pipe(z.enum(["%_of_weight", "target_rpe", ""]))
    .optional(),
  weekly_rpe: z.array(rpeSchema).optional(),
  weekly_percentage: z.array(percentageSchema).optional(),
});

const setGroupSchema = z.object({
  group_id: idSchema,
  variable_parameter: z
    .string()
    .transform((str) => str.trim())
    .transform((str) => str.toLowerCase())
    .pipe(z.enum(["reps", "weight", "rpe"])),
  weekly_notes: z.array(z.string().transform(safeSanitizeNote)),
  weekly_num_sets: limitedArray(z.number().int().min(0).max(15)),
  weekly_reps: z.array(repsSchema),
  weekly_rpe: z.array(rpeSchema),
  weekly_weight_percentage: z.array(percentageSchema),
  mix_weight_config: mixWeightConfigSchema,
  backoff_config: backoffConfigSchema,
  fatigue_drop_config: fatigueDropConfigSchema,
});

const oneRmAnchorSchema = z.object({
  enabled: z.boolean(),
  lift_type: anchorLiftSchema.optional(),
  ratio: anchorRatioSchema.optional(),
});

const deloadConfigSchema = z.object({
  enabled: z.boolean(),
  percentage: percentageSchema.optional(),
});

export const exerciseSchema = z.object({
  _id: idSchema,
  program_id: idSchema,
  block_id: idSchema,
  exercise_name: z.string().transform(safeSanitizeName),
  one_rm_anchor: oneRmAnchorSchema,
  deload_config: deloadConfigSchema.optional(),
  set_groups: limitedArray(setGroupSchema, 15),
  created_at: isoDateTimeSchema.optional(),
  updated_at: isoDateTimeSchema.optional(),
});

const exerciseSetSchema = z.object({
  weight: optional(weightSchema),
  reps: optional(repsSchema),
  rpe: optional(rpeSchema),
  completed: z.boolean(),
});

const exerciseWeeklySchema = z.object({
  sets: z.array(exerciseSetSchema),
  note: z.string().transform(safeSanitizeNote),
});

export const exerciseRecordingSchema = z.object({
  exercise_id: idSchema,
  weekly: z.array(exerciseWeeklySchema),
});

export const validateExercise = (data) => {
  try {
    const result = exerciseSchema.parse(data);
    return { success: true, data: result };
  } catch (error) {
    const formattedErrors = formatValidationErrors(error);
    return {
      success: false,
      error: formattedErrors,
      details: error.format ? error.format() : null,
    };
  }
};
