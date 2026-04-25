import { z } from "zod";
import { FILE_SIZE_LIMITS } from "../importLimits.js";

const ISO_DATETIME_REGEX =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:?\d{2})$/;
const ISO_DATE_REGEX = /^\d{4}-\d{2}-\d{2}$/;
const MAX_NAME_LENGTH = 80;

const sanitizeString = (str) => str.trim();

export const idSchema = z
  .string()
  .min(1, "ID cannot be empty")
  .max(200, "ID too long")
  .regex(
    /^[a-zA-Z0-9_-]+$/,
    "ID contains invalid characters - only letters, numbers, underscore and dash allowed"
  )
  .refine(
    (id) => !id.includes(".."),
    "ID cannot contain path traversal sequences"
  )
  .refine(
    (id) => !id.startsWith("/") && !id.startsWith("\\"),
    "ID cannot start with path separators"
  );

export const safeNameSchema = z
  .string()
  .min(1, "Name cannot be empty")
  .max(MAX_NAME_LENGTH, "Name too long")
  .transform(sanitizeString)
  .refine((str) => str.length > 0, "Name cannot be empty after sanitization");

export const safeDescriptionSchema = z
  .string()
  .max(1000, "Description too long")
  .transform(sanitizeString)
  .optional();

export const safeNoteSchema = z
  .string()
  .max(500, "Note too long")
  .transform(sanitizeString);

export const weightSchema = z
  .number()
  .refine(
    (val) => val === -1 || (val >= 0 && val <= 2000),
    "Weight must be -1 (unset) or between 0-2000"
  )
  .refine((val) => Number.isFinite(val), "Weight must be a valid number");

export const weightDeltaSchema = z
  .number()
  .refine(
    (val) => val === -1 || (val >= -2000 && val <= 2000),
    "Weight delta must be -1 (unset) or between -2000 and 2000"
  )
  .refine((val) => Number.isFinite(val), "Weight delta must be a valid number");

export const repsSchema = z
  .number()
  .int("Reps must be a whole number")
  .refine(
    (val) => val === -1 || (val >= 0 && val <= 50),
    "Reps must be -1 (unset) or between 0-50"
  );

export const rpeSchema = z
  .number()
  .refine(
    (val) => val === -1 || (val >= 1 && val <= 10),
    "RPE must be -1 (unset) or between 1-10"
  )
  .refine((val) => Number.isFinite(val), "RPE must be a valid number");

export const percentageSchema = z
  .number()
  .refine(
    (val) => val === -1 || (val >= 0 && val <= 200),
    "Percentage must be -1 (unset) or between 0-200%"
  )
  .refine((val) => Number.isFinite(val), "Percentage must be a valid number");

export const limitedArray = (schema, maxItems = 15) =>
  z.array(schema).max(maxItems, `Array exceeds maximum of ${maxItems} items`);

export const optional = (schema) =>
  z
    .union([z.literal(-1), z.literal(""), z.undefined(), z.null(), schema])
    .optional();

export const anchorRatioSchema = z
  .number()
  .refine(
    (val) => val === -1 || (val >= 0 && val <= 150),
    "Anchor ratio must be -1 (unset) or between 1-150%"
  )
  .refine((val) => Number.isFinite(val), "Anchor ratio must be a valid number");

export const anchorLiftSchema = z
  .string()
  .min(1, "Anchor lift name cannot be empty")
  .max(20, "Anchor lift name cannot exceed 20 characters")
  .transform(sanitizeString);

export const customAnchoredLiftNameSchema = z
  .string()
  .transform((str) => str.trim())
  .refine((str) => str.length > 0, "Lift name cannot be empty after trimming")
  .refine(
    (str) => str.length <= 20,
    "Lift name cannot exceed 20 characters"
  )
  .refine(
    (str) => !["squat", "bench", "deadlift"].includes(str.toLowerCase()),
    "Cannot use reserved lift names: squat, bench, deadlift"
  );

export const validateCustomLiftName = (liftName, existingLifts = []) => {
  const trimmed = liftName.trim();

  try {
    customAnchoredLiftNameSchema.parse(liftName);
  } catch (error) {
    return { valid: false, error: error.errors[0].message };
  }

  const lowerTrimmed = trimmed.toLowerCase();
  const isDuplicate = existingLifts.some(
    (existing) => existing.toLowerCase() === lowerTrimmed
  );
  if (isDuplicate) {
    return { valid: false, error: `Lift '${trimmed}' already exists` };
  }

  const totalLifts = 3 + existingLifts.length;
  if (totalLifts >= 10) {
    return {
      valid: false,
      error: "Cannot exceed 10 total lifts per program (including squat/bench/deadlift)",
    };
  }

  return { valid: true, trimmed };
};

export const isoDateTimeSchema = z.string().regex(ISO_DATETIME_REGEX).max(50);
export const isoDateSchema = z.string().regex(ISO_DATE_REGEX).max(50);

export const secureObjectValidation = (data) => {
  if (!data || typeof data !== "object") {
    return { valid: false, error: "Invalid data type" };
  }

  const checkCircularReferences = (obj, seen = new WeakSet()) => {
    if (typeof obj === "object" && obj !== null) {
      if (seen.has(obj)) {
        return { valid: false, error: "Circular reference detected" };
      }
      seen.add(obj);

      for (const key in obj) {
        const result = checkCircularReferences(obj[key], seen);
        if (!result.valid) {
          return result;
        }
      }

      seen.delete(obj);
    }
    return { valid: true };
  };

  const circularCheck = checkCircularReferences(data);
  if (!circularCheck.valid) {
    return circularCheck;
  }

  const jsonString = JSON.stringify(data);
  if (jsonString.length > FILE_SIZE_LIMITS.JSON) {
    return {
      valid: false,
      error: `Data size exceeds ${FILE_SIZE_LIMITS.JSON / (1024 * 1024)}MB limit`,
    };
  }

  return { valid: true };
};

export const formatValidationErrors = (zodError) =>
  zodError.errors.map((error) => {
    const pathString = error.path.length > 0 ? error.path.join(".") : "root";
    return {
      path: pathString,
      field: error.path[error.path.length - 1] || "root",
      message: error.message,
      code: error.code,
      expected: error.expected,
      received: error.received,
      receivedValue: JSON.stringify(error.received),
    };
  });

export const safeSanitizeName = (name) => {
  try {
    return safeNameSchema.parse(name);
  } catch (_error) {
    const sanitized = sanitizeString(name || "");
    return sanitized.length > MAX_NAME_LENGTH
      ? sanitized.substring(0, MAX_NAME_LENGTH)
      : sanitized;
  }
};

export const safeSanitizeDescription = (description) => {
  try {
    return safeDescriptionSchema.parse(description);
  } catch (_error) {
    const sanitized = sanitizeString(description || "");
    return sanitized.length > 1000 ? sanitized.substring(0, 1000) : sanitized;
  }
};

export const safeSanitizeNote = (note) => {
  try {
    return safeNoteSchema.parse(note);
  } catch (_error) {
    const sanitized = sanitizeString(note || "");
    return sanitized.length > 500 ? sanitized.substring(0, 500) : sanitized;
  }
};
