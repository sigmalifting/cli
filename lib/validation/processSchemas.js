import { z } from "zod";
import {
  idSchema,
  isoDateTimeSchema,
  isoDateSchema,
  weightSchema,
  optional,
  formatValidationErrors,
  safeSanitizeName,
} from "./shared.js";
import { exerciseRecordingSchema, exerciseSchema } from "./exerciseSchemas.js";
import { programSchema } from "./programSchemas.js";

const processConfigSchema = z.object({
  weight_unit: z
    .string()
    .transform((str) => str.trim())
    .transform((str) => str.toLowerCase())
    .pipe(z.enum(["kg", "lbs"]))
    .optional(),
  weight_rounding: z.number().min(0.1).max(50).optional(),
  week_start_day: z.number().int().min(0).max(6).optional(),
});

const blocklyOneRmSchema = z
  .object({
    squat: optional(weightSchema),
    bench: optional(weightSchema),
    deadlift: optional(weightSchema),
  })
  .catchall(optional(weightSchema));

const oneRmProfileSchema = z
  .object({
    enable_blockly_one_rm: z.boolean(),
    squat: optional(weightSchema),
    bench: optional(weightSchema),
    deadlift: optional(weightSchema),
    blockly_one_rm: z.array(blocklyOneRmSchema).optional(),
  })
  .catchall(optional(weightSchema));

export const processSchema = z.object({
  _id: idSchema,
  name: z.string().transform(safeSanitizeName),
  program_id: idSchema,
  program_name: z.string().transform(safeSanitizeName),
  user_id: optional(idSchema),
  created_at: isoDateTimeSchema,
  updated_at: isoDateTimeSchema,
  start_date: isoDateSchema,
  config: processConfigSchema,
  one_rm_profile: oneRmProfileSchema,
  exercise_recordings: z.array(exerciseRecordingSchema).optional(),
});

const metadataSchema = z
  .object({
    exportedAt: isoDateTimeSchema,
    version: z.string().transform(safeSanitizeName),
    appName: z.string().transform(safeSanitizeName),
  })
  .optional();

export const processImportSchema = z.object({
  process: processSchema,
  program: programSchema,
  exercises: z.array(exerciseSchema),
  metadata: metadataSchema,
});

export const validateProcessImport = (data) => {
  try {
    const result = processImportSchema.parse(data);
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
