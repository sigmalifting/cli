import { z } from "zod";
import {
  idSchema,
  isoDateTimeSchema,
  safeNameSchema,
  limitedArray,
  formatValidationErrors,
  safeSanitizeName,
  safeSanitizeDescription,
  customAnchoredLiftNameSchema,
} from "./shared.js";
import { exerciseSchema } from "./exerciseSchemas.js";

const trainingDaySchema = z.object({
  _id: idSchema,
  name: z.string().transform(safeSanitizeName),
  exercises: z.array(idSchema),
});

const blockSchema = z.object({
  _id: idSchema,
  name: z.string().transform(safeSanitizeName),
  program_id: idSchema,
  duration: z.number().int().min(1).max(15),
  weekly_schedule: z.array(z.string().max(200)).length(7),
  training_days: limitedArray(trainingDaySchema, 100),
});

export const programSchema = z.object({
  _id: idSchema,
  name: z.string().transform(safeSanitizeName),
  description: z.string().transform(safeSanitizeDescription),
  created_at: isoDateTimeSchema,
  updated_at: isoDateTimeSchema,
  blocks: limitedArray(blockSchema, 15),
  custom_anchored_lifts: limitedArray(customAnchoredLiftNameSchema, 10).default([]),
});

const metadataSchema = z
  .object({
    exportedAt: isoDateTimeSchema,
    version: safeNameSchema,
    appName: safeNameSchema,
  })
  .optional();

export const programImportSchema = z.object({
  program: programSchema,
  exercises: z.array(exerciseSchema),
  metadata: metadataSchema,
});

export const validateProgramImport = (data) => {
  try {
    const result = programImportSchema.parse(data);
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
