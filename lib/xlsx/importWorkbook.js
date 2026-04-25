import ExcelJS from "exceljs";
import { TABLE_CONSTANTS } from "./constants.js";
import { generateId } from "../idGenerator.js";
import { FILE_SIZE_LIMITS } from "../importLimits.js";

const debugLog = () => {};
const warnLog = () => {};
const errorLog = () => {};

// Configuration constants
const PARSE_CONFIG = {
  MAX_FILE_SIZE: FILE_SIZE_LIMITS.EXCEL,
  DEFAULT_WEIGHT_UNIT: "lbs",
  DEFAULT_WEIGHT_ROUNDING: 2.5,
  DEFAULT_WEEK_START: 0,
  DEFAULT_FILL_VALUE: -1,
  MAX_SEARCH_ROWS: 1000,
  MAX_NAME_LENGTH: 70,
  MAX_ID_LENGTH: 200,
  MAX_ATTRIBUTE_CACHE_SIZE: 300, // Limit attribute cache to prevent memory issues
  MAX_ROW_CACHE_SIZE: 500, // Limit row cache per worksheet
  PREFIXES: {
    DAY: "$ ",
    EXERCISE: "# ",
    WEEKLY_SCHEDULE: "-> Weekly Schedule",
  },
  PATTERNS: {
    WEEK_HEADER: /Week (\d+)/,
    DAY_HEADER: /Day (\d+)/,
  },
};

// Pre-compiled regex patterns for better performance
const COMPILED_PATTERNS = {
  backoffPercent: /backoff\s+(\d+(?:\.\d+)?)%\s+from\s+(set\s+\d+|last\s+set)/i,
  backoffRpe: /backoff\s+to\s+RPE\s+([\d.]+)\s+from\s+(set\s+\d+|last\s+set)/i,
  fatiguePercent:
    /fatigue\s+drop\s+(\d+(?:\.\d+)?)%\s+with\s+RPE\s+cap\s+([\d.]+)/i,
  fatigueRpe:
    /fatigue\s+drop\s+to\s+RPE\s+([\d.]+)\s+with\s+RPE\s+cap\s+([\d.]+)/i,
  weightPercentage: /percentage:\s*(\d+(?:\.\d+)?)%/i,
  mixedWeight:
    /percentage:\s*(\d+(?:\.\d+)?)%\s+absolute:\s*(-?\d+(?:\.\d+)?)(kg|lbs)/i,
  deloadPercentage: /deload\s+(\d+(?:\.\d+)?)%\s+from\s+week\s+\d+/i,
  setNumber: /set\s+(\d+)/i,
  mergeRange: /([A-Z]+)(\d+)/,
};

// Attribute parsing patterns with pre-compiled regex
const ATTRIBUTE_PATTERNS = {
  backoffPercent: {
    regex: COMPILED_PATTERNS.backoffPercent,
    handler: (match) => ({
      enabled: true,
      type: "%_of_weight",
      percentage: parseFloat(match[1]),
      dependsOn: match[2],
    }),
  },
  backoffRpe: {
    regex: COMPILED_PATTERNS.backoffRpe,
    handler: (match) => ({
      enabled: true,
      type: "target_rpe",
      targetRpe: parseFloat(match[1]),
      dependsOn: match[2],
    }),
  },
  fatiguePercent: {
    regex: COMPILED_PATTERNS.fatiguePercent,
    handler: (match) => ({
      enabled: true,
      type: "%_of_weight",
      percentage: parseFloat(match[1]),
      rpeCap: parseFloat(match[2]),
    }),
  },
  fatigueRpe: {
    regex: COMPILED_PATTERNS.fatigueRpe,
    handler: (match) => ({
      enabled: true,
      type: "target_rpe",
      targetRpe: parseFloat(match[1]),
      rpeCap: parseFloat(match[2]),
    }),
  },
  weightPercentage: {
    regex: COMPILED_PATTERNS.weightPercentage,
    handler: (match) => parseFloat(match[1]),
  },
  mixedWeight: {
    regex: COMPILED_PATTERNS.mixedWeight,
    handler: (match) => ({
      percentage: parseFloat(match[1]),
      absolute: parseFloat(match[2]),
      unit: match[3],
    }),
  },
  deloadPercentage: {
    regex: COMPILED_PATTERNS.deloadPercentage,
    handler: (match) => parseFloat(match[1]),
  },
};

// Cache for worksheet data to avoid repeated access
class WorksheetCache {
  constructor(worksheet) {
    this.worksheet = worksheet;
    this.rowCache = new Map();
    this.merges = this.parseMerges();
  }

  getRow(rowNum) {
    if (!this.rowCache.has(rowNum)) {
      // Check cache size limit before adding
      if (this.rowCache.size >= PARSE_CONFIG.MAX_ROW_CACHE_SIZE) {
        // Clear cache when limit reached
        this.rowCache.clear();
      }

      const row = this.worksheet.getRow(rowNum);
      const values = [];
      for (let i = 1; i <= row.cellCount; i++) {
        values[i] = row.getCell(i).value;
      }
      this.rowCache.set(rowNum, values);
    }
    return this.rowCache.get(rowNum);
  }

  getCellValue(row, col) {
    const rowData = this.getRow(row);
    return rowData ? rowData[col] : null;
  }

  parseMerges() {
    const merges = new Map();
    const worksheetMerges = this.worksheet.model.merges || [];

    for (const merge of worksheetMerges) {
      const [start, end] = merge.split(":");
      const startMatch = start.match(COMPILED_PATTERNS.mergeRange);
      const endMatch = end.match(COMPILED_PATTERNS.mergeRange);

      if (startMatch && endMatch) {
        const startRow = parseInt(startMatch[2]);
        const endRow = parseInt(endMatch[2]);
        const startCol = this.worksheet.getColumn(startMatch[1]).number;
        const endCol = this.worksheet.getColumn(endMatch[1]).number;

        // Store merge info by starting cell
        const key = `${startRow},${startCol}`;
        merges.set(key, {
          startRow,
          endRow,
          startCol,
          endCol,
          span: endCol - startCol + 1,
        });
      }
    }
    return merges;
  }

  isMergedCell(row, col, expectedSpan) {
    // Check if this cell is part of a merge with the expected span
    for (let c = col; c >= Math.max(1, col - expectedSpan + 1); c--) {
      const key = `${row},${c}`;
      const merge = this.merges.get(key);
      if (
        merge &&
        col >= merge.startCol &&
        col <= merge.endCol &&
        merge.span === expectedSpan
      ) {
        return true;
      }
    }
    return false;
  }
}

// Helper functions
const createWeeklyArray = (
  weeks,
  defaultValue = PARSE_CONFIG.DEFAULT_FILL_VALUE
) => Array(weeks).fill(defaultValue);

const wrapError = (operation, errorMessage) => {
  try {
    return operation();
  } catch (error) {
    errorLog(`${errorMessage}:`, error);
    throw new Error(`${errorMessage}: ${error.message}`);
  }
};

/**
 * Generate unique ID combining name and context
 */
function generateContextualId(name) {
  const randomPart = generateId();
  const sanitizedName = name.toLowerCase().replace(/[^a-z0-9]/g, "_");
  const truncatedName =
    sanitizedName.length > PARSE_CONFIG.MAX_NAME_LENGTH
      ? sanitizedName.substring(0, PARSE_CONFIG.MAX_NAME_LENGTH)
      : sanitizedName;

  return `${truncatedName}_${randomPart}`;
}

/**
 * Fast boundary finder using cached worksheet data
 */
function findBoundariesFast(
  cache,
  startCol,
  prefix,
  expectedSpan,
  parser,
  maxRow = PARSE_CONFIG.MAX_SEARCH_ROWS
) {
  const boundaries = [];
  const rowCount = Math.min(cache.worksheet.rowCount || maxRow, maxRow);

  for (let row = 1; row <= rowCount; row++) {
    const value = cache.getCellValue(row, startCol);
    if (
      value &&
      typeof value === "string" &&
      value.startsWith(prefix) &&
      (!expectedSpan || cache.isMergedCell(row, startCol, expectedSpan))
    ) {
      boundaries.push({
        row,
        headerText: value,
        ...(parser ? parser(value) : {}),
      });
    }
  }

  return boundaries;
}

/**
 * Find day boundaries in worksheet
 */
function findDayBoundaries(
  cache,
  startCol,
  maxRow = PARSE_CONFIG.MAX_SEARCH_ROWS
) {
  return findBoundariesFast(
    cache,
    startCol,
    PARSE_CONFIG.PREFIXES.DAY,
    TABLE_CONSTANTS.COLS_PER_DAY,
    (headerText) => {
      const parts = headerText.split(" - ");
      const dayName = parts[parts.length - 1];
      const weekNum =
        parts[0].match(PARSE_CONFIG.PATTERNS.WEEK_HEADER)?.[1] || "1";
      const dayNum =
        parts[1].match(PARSE_CONFIG.PATTERNS.DAY_HEADER)?.[1] || "1";

      return {
        dayName,
        weekNum: parseInt(weekNum),
        dayNum: parseInt(dayNum),
      };
    },
    maxRow
  );
}

/**
 * Find exercise boundaries within a day
 */
function findExerciseBoundaries(cache, startRow, endRow, startCol) {
  const boundaries = [];

  for (let row = startRow; row <= endRow; row++) {
    const value = cache.getCellValue(row, startCol);
    if (
      value &&
      typeof value === "string" &&
      value.startsWith(PARSE_CONFIG.PREFIXES.EXERCISE) &&
      cache.isMergedCell(row, startCol, TABLE_CONSTANTS.COLS_PER_DAY)
    ) {
      boundaries.push({
        row,
        headerText: value,
      });
    }
  }

  return boundaries;
}

/**
 * Parse attributes with memoization
 */
const attributeCache = new Map();

function parseAttributesAdvanced(attributes) {
  if (!attributes) {
    return {
      backoffConfig: { enabled: false },
      fatigueDropConfig: { enabled: false },
      weightPercentage: null,
      mixedWeight: null,
      deloadPercentage: null,
    };
  }

  // Check cache first
  if (attributeCache.has(attributes)) {
    return attributeCache.get(attributes);
  }

  // Check cache size limit before adding
  if (attributeCache.size >= PARSE_CONFIG.MAX_ATTRIBUTE_CACHE_SIZE) {
    // Clear cache when limit reached
    attributeCache.clear();
  }

  const result = {
    backoffConfig: { enabled: false },
    fatigueDropConfig: { enabled: false },
    weightPercentage: null,
    mixedWeight: null,
    deloadPercentage: null,
  };

  const trimmed = attributes.trim();

  // Check patterns in order of likelihood
  const checks = [
    ["backoffPercent", "backoffConfig"],
    ["backoffRpe", "backoffConfig"],
    ["fatiguePercent", "fatigueDropConfig"],
    ["fatigueRpe", "fatigueDropConfig"],
    ["mixedWeight", "mixedWeight"], // Check mixed weight before regular percentage
    ["weightPercentage", "weightPercentage"],
    ["deloadPercentage", "deloadPercentage"],
  ];

  for (const [patternName, resultKey] of checks) {
    const pattern = ATTRIBUTE_PATTERNS[patternName];
    const match = trimmed.match(pattern.regex);
    if (match) {
      const value = pattern.handler(match);
      if (resultKey === "backoffConfig" || resultKey === "fatigueDropConfig") {
        result[resultKey] = value;
      } else {
        result[resultKey] = value;
      }
    }
  }

  // Cache the result
  attributeCache.set(attributes, result);
  return result;
}

/**
 * Parse set data with cached worksheet access
 */
function parseSetDataAdvanced(cache, row, startCol) {
  const { PRESC_WEIGHT, PRESC_REPS, PRESC_RPE, PRESC_ATTRIBUTES, PRESC_NOTES } =
    TABLE_CONSTANTS.DAY_COLS;

  const rowData = cache.getRow(row);
  const weight = rowData[startCol + PRESC_WEIGHT];
  const reps = rowData[startCol + PRESC_REPS];
  const rpe = rowData[startCol + PRESC_RPE];
  const attributes = rowData[startCol + PRESC_ATTRIBUTES] || "";
  const notes = rowData[startCol + PRESC_NOTES] || "";

  const hasWeight = weight !== null && weight !== undefined && weight !== "";
  const hasReps = reps !== null && reps !== undefined && reps !== "";
  const hasRpe = rpe !== null && rpe !== undefined && rpe !== "";

  let variableParameter = "rpe";

  if (!hasWeight && hasReps && hasRpe) {
    variableParameter = "weight";
  } else if (hasWeight && !hasReps && hasRpe) {
    variableParameter = "reps";
  } else if (hasWeight && hasReps && !hasRpe) {
    variableParameter = "rpe";
  } else if (!hasWeight && !hasReps && hasRpe) {
    variableParameter = "weight";
  } else if (!hasWeight && hasReps && !hasRpe) {
    variableParameter = "rpe";
  }

  return {
    weight: hasWeight ? parseFloat(weight) : null,
    reps: hasReps ? parseFloat(reps) : null,
    rpe: hasRpe ? parseFloat(rpe) : null,
    attributes,
    notes,
    variableParameter,
  };
}

/**
 * Batch process exercise data for better performance
 */
function parseExerciseDataWithBoundariesFast(
  cache,
  startRow,
  endRow,
  startCol,
  parsed,
  totalWeeks
) {
  const { SET_NO } = TABLE_CONSTANTS.DAY_COLS;
  const tempSetGroups = [];
  let currentRow = startRow + 1;
  let currentSetGroup = 0;
  let setsInCurrentGroup = 0;
  const setGroupSizes = parsed.setGroupSizes;

  // Batch read all rows
  const rows = [];
  for (let r = currentRow; r <= endRow; r++) {
    const setNoValue = cache.getCellValue(r, startCol + SET_NO);
    if (!setNoValue || typeof setNoValue !== "number") break;
    rows.push(r);
  }

  // Process rows
  for (const row of rows) {
    if (
      currentSetGroup < setGroupSizes.length &&
      setsInCurrentGroup >= setGroupSizes[currentSetGroup]
    ) {
      currentSetGroup++;
      setsInCurrentGroup = 0;
    }

    if (setsInCurrentGroup === 0) {
      const setData = parseSetDataAdvanced(cache, row, startCol);
      const parsedAttrs = parseAttributesAdvanced(setData.attributes);

      tempSetGroups.push({
        groupId: generateId(),
        groupIndex: currentSetGroup,
        numSets: setGroupSizes[currentSetGroup] || 1,
        setData,
        parsedAttributes: parsedAttrs,
      });
    }

    setsInCurrentGroup++;
  }

  // CRITICAL: Handle case where metadata indicates N set groups but Week 1 has 0 sets
  // If we didn't create enough set groups from actual rows, create placeholders
  // This happens when Week 1 has all zeros: metadata says "set groups: 1" or "set groups: 0+0"
  // but there are no actual set rows to process
  while (tempSetGroups.length < setGroupSizes.length) {
    const groupIndex = tempSetGroups.length;
    tempSetGroups.push({
      groupId: generateId(),
      groupIndex: groupIndex,
      numSets: setGroupSizes[groupIndex] || 0,
      setData: {
        weight: null,
        reps: null,
        rpe: null,
        attributes: "",
        notes: "",
        variableParameter: "rpe",
      },
      parsedAttributes: {
        backoffConfig: { enabled: false },
        fatigueDropConfig: { enabled: false },
        weightPercentage: null,
        mixedWeight: null,
        deloadPercentage: null,
      },
    });
  }

  // Build final set groups
  const setGroups = tempSetGroups.map((tempGroup, index) => {
    const { setData, parsedAttributes } = tempGroup;

    const setGroup = {
      group_id: tempGroup.groupId,
      variable_parameter: setData.variableParameter || "rpe",
      weekly_notes: createWeeklyArray(totalWeeks, ""),
      weekly_num_sets: createWeeklyArray(totalWeeks, tempGroup.numSets),
      weekly_reps: createWeeklyArray(totalWeeks),
      weekly_rpe: createWeeklyArray(totalWeeks),
      weekly_weight_percentage: createWeeklyArray(totalWeeks),
      mix_weight_config: { enabled: false },
      backoff_config: {
        enabled: parsedAttributes.backoffConfig.enabled,
        depends_on_set_group_id: "",
        type: parsedAttributes.backoffConfig.type || "",
      },
      fatigue_drop_config: {
        enabled: parsedAttributes.fatigueDropConfig.enabled,
        type: parsedAttributes.fatigueDropConfig.type || "",
      },
    };

    // Initialize arrays based on config
    if (parsedAttributes.backoffConfig.enabled) {
      if (parsedAttributes.backoffConfig.type === "target_rpe") {
        setGroup.backoff_config.weekly_rpe = createWeeklyArray(totalWeeks);
      } else if (parsedAttributes.backoffConfig.type === "%_of_weight") {
        setGroup.backoff_config.weekly_percentage =
          createWeeklyArray(totalWeeks);
      }
    }

    if (parsedAttributes.fatigueDropConfig.enabled) {
      setGroup.fatigue_drop_config.rpe_cap = createWeeklyArray(totalWeeks);
      if (parsedAttributes.fatigueDropConfig.type === "target_rpe") {
        setGroup.fatigue_drop_config.weekly_rpe = createWeeklyArray(totalWeeks);
      } else if (parsedAttributes.fatigueDropConfig.type === "%_of_weight") {
        setGroup.fatigue_drop_config.weekly_percentage =
          createWeeklyArray(totalWeeks);
      }
    }

    // Set first week values
    if (setData.reps !== null && setGroup.variable_parameter !== "reps") {
      setGroup.weekly_reps[0] = setData.reps;
    }
    if (setData.rpe !== null && setGroup.variable_parameter !== "rpe") {
      setGroup.weekly_rpe[0] = setData.rpe;
    }

    // Handle mixed weight configuration
    if (
      parsedAttributes.mixedWeight !== null &&
      setGroup.variable_parameter !== "weight"
    ) {
      setGroup.mix_weight_config = {
        enabled: true,
        weight_unit: parsedAttributes.mixedWeight.unit,
        weekly_weight_percentage: createWeeklyArray(totalWeeks, -1),
        weekly_weight_absolute: createWeeklyArray(totalWeeks, -1),
      };
      setGroup.mix_weight_config.weekly_weight_percentage[0] =
        parsedAttributes.mixedWeight.percentage;
      setGroup.mix_weight_config.weekly_weight_absolute[0] =
        parsedAttributes.mixedWeight.absolute;
    } else if (
      parsedAttributes.weightPercentage !== null &&
      setGroup.variable_parameter !== "weight"
    ) {
      setGroup.weekly_weight_percentage[0] = parsedAttributes.weightPercentage;
    }

    // Set config values for first week
    if (parsedAttributes.backoffConfig.enabled) {
      if (
        parsedAttributes.backoffConfig.type === "target_rpe" &&
        setGroup.backoff_config.weekly_rpe
      ) {
        setGroup.backoff_config.weekly_rpe[0] =
          parsedAttributes.backoffConfig.targetRpe ||
          PARSE_CONFIG.DEFAULT_FILL_VALUE;
      } else if (
        parsedAttributes.backoffConfig.type === "%_of_weight" &&
        setGroup.backoff_config.weekly_percentage
      ) {
        setGroup.backoff_config.weekly_percentage[0] =
          parsedAttributes.backoffConfig.percentage ||
          PARSE_CONFIG.DEFAULT_FILL_VALUE;
      }
    }

    if (parsedAttributes.fatigueDropConfig.enabled) {
      if (setGroup.fatigue_drop_config.rpe_cap) {
        setGroup.fatigue_drop_config.rpe_cap[0] =
          parsedAttributes.fatigueDropConfig.rpeCap ||
          PARSE_CONFIG.DEFAULT_FILL_VALUE;
      }
      if (
        parsedAttributes.fatigueDropConfig.type === "target_rpe" &&
        setGroup.fatigue_drop_config.weekly_rpe
      ) {
        setGroup.fatigue_drop_config.weekly_rpe[0] =
          parsedAttributes.fatigueDropConfig.targetRpe ||
          PARSE_CONFIG.DEFAULT_FILL_VALUE;
      } else if (
        parsedAttributes.fatigueDropConfig.type === "%_of_weight" &&
        setGroup.fatigue_drop_config.weekly_percentage
      ) {
        setGroup.fatigue_drop_config.weekly_percentage[0] =
          parsedAttributes.fatigueDropConfig.percentage ||
          PARSE_CONFIG.DEFAULT_FILL_VALUE;
      }
    }

    // Handle variable parameter arrays
    if (setGroup.variable_parameter === "weight") {
      setGroup.weekly_weight_percentage = createWeeklyArray(totalWeeks);
    } else if (setGroup.variable_parameter === "reps") {
      setGroup.weekly_reps = createWeeklyArray(totalWeeks);
    } else if (setGroup.variable_parameter === "rpe") {
      setGroup.weekly_rpe = createWeeklyArray(totalWeeks);
    }

    if (setData.notes) {
      setGroup.weekly_notes[0] = setData.notes;
    }

    // Resolve dependencies after all groups are created
    if (
      parsedAttributes.backoffConfig.enabled &&
      parsedAttributes.backoffConfig.dependsOn
    ) {
      setGroup.backoff_config.depends_on_set_group_id =
        resolveBackoffDependency(
          parsedAttributes.backoffConfig.dependsOn,
          tempSetGroups,
          index
        );
    }

    return setGroup;
  });

  return {
    _id: "",
    exercise_name: parsed.name,
    one_rm_anchor: parsed.oneRmAnchor,
    deload_config: {
      ...parsed.deloadConfig,
      percentage: parsed.deloadConfig.percentage || 90,
    },
    set_groups: setGroups,
  };
}

/**
 * Resolve backoff dependency
 */
function resolveBackoffDependency(dependsOn, setGroups, currentGroupIndex) {
  if (!dependsOn || currentGroupIndex === 0) return "";

  const setMatch = dependsOn.match(COMPILED_PATTERNS.setNumber);
  if (setMatch) {
    const targetSetNumber = parseInt(setMatch[1]);
    let setCounter = 0;

    for (let i = 0; i < currentGroupIndex; i++) {
      const groupSize = setGroups[i].numSets || 1;
      if (setCounter + groupSize >= targetSetNumber) {
        return setGroups[i].groupId;
      }
      setCounter += groupSize;
    }
  }

  if (dependsOn.toLowerCase().includes("last set") && currentGroupIndex > 0) {
    return setGroups[currentGroupIndex - 1].groupId;
  }

  return "";
}

/**
 * Fast update of exercise weekly data
 */
function updateExerciseWeeklyDataFast(
  cache,
  exerciseRow,
  weekCol,
  weekIdx,
  exercise,
  knownLifts = ["squat", "bench", "deadlift"]
) {
  // First, check if this week has different set group sizes by parsing the exercise header
  const exerciseHeaderValue = cache.getCellValue(exerciseRow, weekCol);
  if (exerciseHeaderValue && typeof exerciseHeaderValue === "string") {
    const parsed = parseExerciseName(exerciseHeaderValue, knownLifts);

    // Update weekly_num_sets if we found new set group sizes
    if (parsed.setGroupSizes.length > 0) {
      parsed.setGroupSizes.forEach((numSets, groupIndex) => {
        if (groupIndex < exercise.set_groups.length) {
          exercise.set_groups[groupIndex].weekly_num_sets[weekIdx] = numSets;
        }
      });
    }
  }

  let currentRow = exerciseRow + 1;

  exercise.set_groups.forEach((setGroup) => {
    const numSets = setGroup.weekly_num_sets[weekIdx] || 0; // Now uses updated value

    if (numSets > 0) {
      const setData = parseSetDataAdvanced(cache, currentRow, weekCol);
      const parsedAttrs = parseAttributesAdvanced(setData.attributes);

      // CRITICAL: Update variable_parameter if this is the first week with actual data
      // This handles the case where Week 1 had 0 sets (placeholder with default "rpe")
      // but the actual exercise has a different variable parameter
      // Check if all previous weeks had 0 sets
      const hadDataBefore = setGroup.weekly_num_sets
        .slice(0, weekIdx)
        .some(n => n > 0);

      if (!hadDataBefore && setData.variableParameter) {
        // This is the first week with actual data - trust its variable parameter
        setGroup.variable_parameter = setData.variableParameter;

        // Also need to update the corresponding weekly array to be variable
        // and move existing data to the correct array
        if (setData.variableParameter === "weight") {
          setGroup.weekly_weight_percentage = createWeeklyArray(
            setGroup.weekly_num_sets.length
          );
        } else if (setData.variableParameter === "reps") {
          setGroup.weekly_reps = createWeeklyArray(
            setGroup.weekly_num_sets.length
          );
        } else if (setData.variableParameter === "rpe") {
          setGroup.weekly_rpe = createWeeklyArray(
            setGroup.weekly_num_sets.length
          );
        }
      }

      // Update values
      // Handle mixed weight configuration first
      if (
        parsedAttrs.mixedWeight !== null &&
        setGroup.variable_parameter !== "weight"
      ) {
        // Enable mix_weight_config if not already enabled
        if (!setGroup.mix_weight_config.enabled) {
          setGroup.mix_weight_config = {
            enabled: true,
            weight_unit: parsedAttrs.mixedWeight.unit,
            weekly_weight_percentage: createWeeklyArray(
              exercise.set_groups[0].weekly_num_sets.length,
              -1
            ),
            weekly_weight_absolute: createWeeklyArray(
              exercise.set_groups[0].weekly_num_sets.length,
              -1
            ),
          };
        }
        setGroup.mix_weight_config.weekly_weight_percentage[weekIdx] =
          parsedAttrs.mixedWeight.percentage;
        setGroup.mix_weight_config.weekly_weight_absolute[weekIdx] =
          parsedAttrs.mixedWeight.absolute;
      } else if (
        parsedAttrs.weightPercentage !== null &&
        parsedAttrs.weightPercentage !== undefined
      ) {
        setGroup.weekly_weight_percentage[weekIdx] =
          parsedAttrs.weightPercentage;
      } else if (
        setData.weight !== null &&
        setGroup.variable_parameter !== "weight"
      ) {
        setGroup.weekly_weight_percentage[weekIdx] =
          PARSE_CONFIG.DEFAULT_FILL_VALUE;
      }
      if (setData.reps !== null && setGroup.variable_parameter !== "reps") {
        setGroup.weekly_reps[weekIdx] = setData.reps;
      }
      if (setData.rpe !== null && setGroup.variable_parameter !== "rpe") {
        setGroup.weekly_rpe[weekIdx] = setData.rpe;
      }
      if (setData.notes) {
        setGroup.weekly_notes[weekIdx] = setData.notes;
      }

      // Update configs
      if (
        parsedAttrs.backoffConfig.enabled &&
        setGroup.backoff_config.enabled
      ) {
        if (
          parsedAttrs.backoffConfig.type === "%_of_weight" &&
          setGroup.backoff_config.weekly_percentage
        ) {
          setGroup.backoff_config.weekly_percentage[weekIdx] =
            parsedAttrs.backoffConfig.percentage ||
            PARSE_CONFIG.DEFAULT_FILL_VALUE;
        } else if (
          parsedAttrs.backoffConfig.type === "target_rpe" &&
          setGroup.backoff_config.weekly_rpe
        ) {
          setGroup.backoff_config.weekly_rpe[weekIdx] =
            parsedAttrs.backoffConfig.targetRpe ||
            PARSE_CONFIG.DEFAULT_FILL_VALUE;
        }
      }

      if (
        parsedAttrs.fatigueDropConfig.enabled &&
        setGroup.fatigue_drop_config.enabled
      ) {
        if (setGroup.fatigue_drop_config.rpe_cap) {
          setGroup.fatigue_drop_config.rpe_cap[weekIdx] =
            parsedAttrs.fatigueDropConfig.rpeCap ||
            PARSE_CONFIG.DEFAULT_FILL_VALUE;
        }
        if (
          parsedAttrs.fatigueDropConfig.type === "%_of_weight" &&
          setGroup.fatigue_drop_config.weekly_percentage
        ) {
          setGroup.fatigue_drop_config.weekly_percentage[weekIdx] =
            parsedAttrs.fatigueDropConfig.percentage ||
            PARSE_CONFIG.DEFAULT_FILL_VALUE;
        } else if (
          parsedAttrs.fatigueDropConfig.type === "target_rpe" &&
          setGroup.fatigue_drop_config.weekly_rpe
        ) {
          setGroup.fatigue_drop_config.weekly_rpe[weekIdx] =
            parsedAttrs.fatigueDropConfig.targetRpe ||
            PARSE_CONFIG.DEFAULT_FILL_VALUE;
        }
      }
    }

    currentRow += numSets;
  });
}

/**
 * Sanitize workbook by removing all potentially dangerous content
 */
const sanitizeWorkbook = (workbook) => {
  if (!workbook) {
    throw new Error("Invalid workbook object for sanitization");
  }

  try {
    const propsToDelete = ["vbaProject", "calcProperties", "customProperties"];
    propsToDelete.forEach((prop) => delete workbook[prop]);

    if (workbook.views) workbook.views = [];
  } catch (error) {
    throw new Error(`Workbook sanitization failed: ${error.message}`);
  }
};

/**
 * Parse metadata worksheet optimized
 */
function parseMetadata(worksheet) {
  const defaultMetadata = {
    programName: "",
    processName: "",
    programOnly: null,
    exportDate: "",
    appVersion: "1.0", //usage to TBD
  };

  if (!worksheet) return defaultMetadata;

  const metadata = { ...defaultMetadata };
  const cache = new WorksheetCache(worksheet);

  const fieldMapping = {
    "program name": {
      key: "programName",
    },
    "process name": {
      key: "processName",
    },
    "program only": {
      key: "programOnly",
      parse: (value) => value.toLowerCase() === "yes",
    },
    "export date": {
      key: "exportDate",
    },
    "app version": {
      key: "appVersion",
    },
  };

  // Read only first 20 rows
  for (let row = 1; row <= 20; row++) {
    const labelValue = cache.getCellValue(row, 1);
    const dataValue = cache.getCellValue(row, 2);

    if (labelValue && dataValue) {
      const label = labelValue
        .toString()
        .toLowerCase()
        .replace(/:$/, "")
        .trim();
      const value = dataValue.toString();

      for (const [key, config] of Object.entries(fieldMapping)) {
        if (label === key) {
          metadata[config.key] = config.parse ? config.parse(value) : value;
          break;
        }
      }
    }
  }

  return metadata;
}

/**
 * Parse 1RM Profile worksheet optimized
 */
function parseOneRmProfile(worksheet) {
  const defaultProfile = {
    oneRmData: {
      enable_blockly_one_rm: false,
      squat: PARSE_CONFIG.DEFAULT_FILL_VALUE,
      bench: PARSE_CONFIG.DEFAULT_FILL_VALUE,
      deadlift: PARSE_CONFIG.DEFAULT_FILL_VALUE,
      blockly_one_rm: [],
    },
    weightUnit: PARSE_CONFIG.DEFAULT_WEIGHT_UNIT,
    weightRounding: PARSE_CONFIG.DEFAULT_WEIGHT_ROUNDING,
    weekStartDay: PARSE_CONFIG.DEFAULT_WEEK_START,
  };

  if (!worksheet) return defaultProfile;

  const result = { ...defaultProfile };
  const cache = new WorksheetCache(worksheet);
  const dayNames = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];

  // Find configuration values
  for (let row = 1; row <= 15; row++) {
    const label = cache.getCellValue(row, 1);
    const value = cache.getCellValue(row, 2);

    if (label === "Weight Unit:" && value) {
      result.weightUnit = value.toString().toLowerCase();
    } else if (label === "Weight Rounding:" && value && !isNaN(value)) {
      result.weightRounding = parseFloat(value);
    } else if (label === "Day 1 starts on:" && value) {
      const dayIndex = dayNames.findIndex(
        (d) => d.toLowerCase() === value.toString().toLowerCase()
      );
      if (dayIndex !== -1) {
        result.weekStartDay = dayIndex;
      }
    }
  }

  // Find global 1RM section
  let globalSectionRow = 0;
  for (let row = 1; row <= 20; row++) {
    if (cache.getCellValue(row, 1) === "Global 1RM Profile") {
      globalSectionRow = row + 2;
      break;
    }
  }

  // Array to store custom lift names discovered during import
  const customLifts = [];

  if (globalSectionRow > 0) {
    const exercises = ["Squat", "Bench", "Deadlift"];
    const keys = ["squat", "bench", "deadlift"];

    // First, read the reserved lifts (S/B/D)
    exercises.forEach((exercise, idx) => {
      if (cache.getCellValue(globalSectionRow + idx, 1) === exercise) {
        const value = parseFloat(cache.getCellValue(globalSectionRow + idx, 2));
        if (!isNaN(value) && value > 0) {
          result.oneRmData[keys[idx]] = value;
        }
      }
    });

    // Then, scan for custom anchor lifts after the reserved lifts
    const customLiftStartRow = globalSectionRow + 3; // Start after S/B/D
    const maxCustomLifts = 7; // Max 7 custom lifts allowed (spec: 3 reserved + 7 custom = 10 total)

    for (let i = 0; i < maxCustomLifts; i++) {
      const currentRow = customLiftStartRow + i;
      const liftNameCell = cache.getCellValue(currentRow, 1);
      const liftValueCell = cache.getCellValue(currentRow, 2);

      // Stop scanning if we hit an empty row (no lift name)
      if (!liftNameCell || liftNameCell.toString().trim() === "") {
        break;
      }

      const liftName = liftNameCell.toString().trim();

      // Validate: Skip reserved names (should not appear here, but be defensive)
      const reservedNames = ["Squat", "Bench", "Deadlift"];
      if (reservedNames.some(reserved => reserved.toLowerCase() === liftName.toLowerCase())) {
        warnLog(`[Excel Import] Skipping reserved lift name in custom section: ${liftName}`);
        continue;
      }

      // Validate: Enforce length limits (1-20 characters per spec)
      if (liftName.length < 1 || liftName.length > 20) {
        warnLog(`[Excel Import] Skipping invalid lift name (invalid length: ${liftName.length}): ${liftName}`);
        continue;
      }

      // Validate: Check for uniqueness (case-insensitive)
      const isDuplicate = customLifts.some(
        existing => existing.toLowerCase() === liftName.toLowerCase()
      );
      if (isDuplicate) {
        warnLog(`[Excel Import] Skipping duplicate custom lift: ${liftName}`);
        continue;
      }

      // Parse the 1RM value
      const value = parseFloat(liftValueCell);
      // Store with lowercase key in one_rm_profile (consistent with spec D3)
      if (!isNaN(value) && value > 0) {
        result.oneRmData[liftName.toLowerCase()] = value;
      } else {
        // Still add the lift with default -1 value if no valid value found
        result.oneRmData[liftName.toLowerCase()] = PARSE_CONFIG.DEFAULT_FILL_VALUE;
      }

      // Store the lift name with original capitalization for program.custom_anchored_lifts
      customLifts.push(liftName);
      debugLog(`[Excel Import] Found custom anchor lift: ${liftName}`);
    }
  }

  // Find block-specific section
  let blockSectionRow = 0;
  for (let row = globalSectionRow; row <= globalSectionRow + 20; row++) {
    if (cache.getCellValue(row, 1) === "Block-Specific 1RM") {
      blockSectionRow = row;
      break;
    }
  }

  if (blockSectionRow > 0) {
    // Look for the enabled status in the next few rows
    let enabledRow = 0;
    for (let row = blockSectionRow + 1; row <= blockSectionRow + 5; row++) {
      const cellValue = cache.getCellValue(row, 1);
      if (cellValue && cellValue.toString().includes("Block-specific 1RM")) {
        result.oneRmData.enable_blockly_one_rm =
          cache.getCellValue(row, 2) === "Yes";
        enabledRow = row;
        break;
      }
    }

    if (result.oneRmData.enable_blockly_one_rm && enabledRow > 0) {
      // Look for the header row starting with "Block"
      let headerRow = 0;
      for (let row = enabledRow + 1; row <= enabledRow + 10; row++) {
        const col1Value = cache.getCellValue(row, 1);
        if (col1Value && col1Value.toString().toLowerCase() === "block") {
          headerRow = row;
          break;
        }
      }

      // Parse column headers to build lift column map
      if (headerRow > 0) {
        // Build array of lift column mappings [{colIndex: 2, liftName: "squat"}, ...]
        const liftColumns = [];

        // Reserved lifts (columns 2-4)
        liftColumns.push({ colIndex: 2, liftName: "squat" });
        liftColumns.push({ colIndex: 3, liftName: "bench" });
        liftColumns.push({ colIndex: 4, liftName: "deadlift" });

        // Parse custom lift columns (columns 5+)
        // Use the custom lifts discovered in global section as the expected list
        if (customLifts.length > 0) {
          customLifts.forEach((liftName, idx) => {
            const colIndex = 5 + idx; // Custom lifts start at column 5
            const headerValue = cache.getCellValue(headerRow, colIndex);

            // Verify header matches expected lift (with flexibility for weight unit suffix)
            if (headerValue) {
              const headerStr = headerValue.toString().trim();
              // Remove weight unit suffix like " (kg)" or " (lbs)"
              const parsedLiftName = headerStr.replace(/\s*\((kg|lbs)\)$/i, "").trim();

              // Match against expected custom lift name (case-insensitive for flexibility)
              if (parsedLiftName.toLowerCase() === liftName.toLowerCase()) {
                liftColumns.push({ colIndex, liftName });
                debugLog(`[Excel Import] Mapped column ${colIndex} to custom lift: ${liftName}`);
              } else {
                // Header mismatch - use position-based mapping (global section is source of truth)
                warnLog(
                  `[Excel Import] Block header mismatch at column ${colIndex}. Expected: "${liftName}", Found: "${parsedLiftName}". Using position-based mapping from global section.`
                );
                // Use expected lift name from global section to maintain data consistency
                liftColumns.push({ colIndex, liftName });
              }
            } else {
              // Missing column header - use position-based mapping, will fill with -1 if no data
              warnLog(
                `[Excel Import] Missing header for custom lift "${liftName}" at column ${colIndex}. Using position-based mapping from global section.`
              );
              liftColumns.push({ colIndex, liftName });
            }
          });
        }

        // Read block data rows using the column map
        // Support up to 20 blocks
        for (let row = headerRow + 1; row <= headerRow + 20; row++) {
          const cellValue = cache.getCellValue(row, 1);
          if (cellValue && cellValue.toString().startsWith("Block ")) {
            const blockData = {};

            // Extract values for all lifts based on column map
            liftColumns.forEach(({ colIndex, liftName }) => {
              const value = parseFloat(cache.getCellValue(row, colIndex));
              // Store with lowercase key in one_rm_profile (consistent with spec D3)
              blockData[liftName.toLowerCase()] =
                !isNaN(value) && value > 0
                  ? value
                  : PARSE_CONFIG.DEFAULT_FILL_VALUE;
            });

            result.oneRmData.blockly_one_rm.push(blockData);
          }
        }
      }
    }
  }

  // Include custom lifts in the result for use in program structure
  result.customLifts = customLifts;

  return result;
}

/**
 * Check if workbook contains execution data - optimized
 */
function checkForExecutionData(blockSheets) {
  const { EXEC_WEIGHT, EXEC_REPS } = TABLE_CONSTANTS.DAY_COLS;

  for (const worksheet of blockSheets) {
    const cache = new WorksheetCache(worksheet);

    // Quick scan first 3 exercises
    let exercisesFound = 0;
    for (let row = 1; row <= 100 && exercisesFound < 3; row++) {
      const cellValue = cache.getCellValue(row, 1);
      if (
        cellValue &&
        typeof cellValue === "string" &&
        cellValue.startsWith(PARSE_CONFIG.PREFIXES.EXERCISE)
      ) {
        exercisesFound++;

        // Check next row for execution data
        const nextRow = row + 1;
        if (
          cache.getCellValue(nextRow, 1 + EXEC_WEIGHT) ||
          cache.getCellValue(nextRow, 1 + EXEC_REPS)
        ) {
          return true;
        }
      }
    }
  }

  return false;
}

/**
 * Parse exercise name and extract metadata
 * @param {string} fullText - Full exercise header text (e.g., "# Overhead Press (max: 90% overhead press, ...)")
 * @param {string[]} knownLifts - Array of known lift names from 1RM Profile (lowercase)
 */
function parseExerciseName(fullText, knownLifts = ["squat", "bench", "deadlift"]) {
  const text = fullText.substring(2).trim();
  // The ? makes it non-greedy so it stops at the last valid metadata parentheses
  //   const parenMatch = text.match(/^(.+?)\s*\(([^)]+)\)$/);
  const parenMatch = text.match(/^(.*)\s*\(([^)]+)\)$/);

  const name = parenMatch ? parenMatch[1].trim() : text;

  let metadata = parenMatch ? parenMatch[2] : "";

  const result = {
    name,
    oneRmAnchor: { enabled: false },
    deloadConfig: { enabled: false },
    setGroupSizes: [],
  };

  if (metadata) {
    // IMPORTANT: Extract the anchor section first before splitting by comma
    // This allows lift names to contain commas without breaking the parser

    // Try to find and extract the anchor section using known lift names
    // Match "max: X% " and capture everything after until we can match a known lift
    const anchorStartMatch = metadata.match(/max:\s*(\d+)%\s*/i);

    if (anchorStartMatch) {
      const percentage = parseInt(anchorStartMatch[1]);
      const afterPercentage = metadata.substring(anchorStartMatch.index + anchorStartMatch[0].length);

      // Find which known lift this references
      // knownLifts is already sorted by length (longest first) for correct matching
      for (const knownLift of knownLifts) {
        // Check if the afterPercentage string starts with this known lift (case-insensitive)
        if (afterPercentage.toLowerCase().startsWith(knownLift.toLowerCase())) {
          result.oneRmAnchor = {
            enabled: true,
            ratio: percentage / 100,
            lift_type: knownLift.toLowerCase(),
          };

          // Remove the entire anchor section from metadata
          // Section is from "max:" to the end of the lift name
          const anchorStart = anchorStartMatch.index;
          const liftEnd = anchorStart + anchorStartMatch[0].length + knownLift.length;

          const beforeAnchor = metadata.substring(0, anchorStart);
          const afterLiftName = metadata.substring(liftEnd);

          // Reconstruct metadata without the anchor section
          metadata = beforeAnchor + afterLiftName;
          // Clean up any leading/trailing commas or whitespace
          metadata = metadata.replace(/^[,\s]+|[,\s]+$/g, '').replace(/,\s*,/g, ',');

          break;
        }
      }
    }

    // Now split by comma for remaining attributes
    const parts = metadata.split(",").map((p) => p.trim()).filter(p => p.length > 0);

    for (const part of parts) {
      const deloadMatch = part.match(/deload week:\s*(\d+)%/);
      if (deloadMatch) {
        result.deloadConfig = {
          enabled: true,
          percentage: parseInt(deloadMatch[1]),
        };
        continue;
      }

      const setGroupMatch = part.match(/set groups:\s*([\d+]+)/);
      if (setGroupMatch) {
        result.setGroupSizes = setGroupMatch[1]
          .split("+")
          .map((n) => parseInt(n));
      }
    }
  }

  if (result.setGroupSizes.length === 0) {
    result.setGroupSizes = [1];
  }

  return result;
}

/**
 * Find and parse weekly schedule section optimized
 */
function findAndParseWeeklySchedule(cache) {
  let scheduleStartCol = 0;
  let scheduleStartRow = 0;

  // Quick scan for schedule marker
  const maxCol = Math.min(cache.worksheet.columnCount || 500, 500);
  for (let col = 1; col <= maxCol; col++) {
    for (let row = 1; row <= 10; row++) {
      const value = cache.getCellValue(row, col);
      if (
        value &&
        typeof value === "string" &&
        value.includes(PARSE_CONFIG.PREFIXES.WEEKLY_SCHEDULE)
      ) {
        scheduleStartCol = col;
        scheduleStartRow = row;
        break;
      }
    }
    if (scheduleStartCol > 0) break;
  }

  if (scheduleStartCol === 0) {
    return createWeeklyArray(7, "");
  }

  const weekly_schedule = createWeeklyArray(7, "");
  const dataRow = scheduleStartRow + 1;

  for (let i = 0; i < 7; i++) {
    const trainingDayValue = cache.getCellValue(
      dataRow + i,
      scheduleStartCol + 1
    );
    if (trainingDayValue && trainingDayValue.toString().trim() !== "") {
      const value = trainingDayValue.toString().trim();
      weekly_schedule[i] = value.includes(" - ")
        ? value.split(" - ").slice(1).join(" - ")
        : value;
    }
  }

  return weekly_schedule;
}

/**
 * Find week header optimized
 */
function findWeekHeader(cache, startCol) {
  for (let row = 1; row <= 10; row++) {
    const value = cache.getCellValue(row, startCol);
    if (value && typeof value === "string" && value.includes("Week ")) {
      return { row, col: startCol, text: value };
    }
  }
  return null;
}

/**
 * Parse a block worksheet optimized
 * @param {object} worksheet - ExcelJS worksheet
 * @param {number} blockIndex - Block index
 * @param {number} weekStartDay - Week start day (0-6)
 * @param {string[]} knownLifts - Known lift names from 1RM Profile (lowercase)
 */
function parseBlockWorksheet(worksheet, blockIndex, weekStartDay = 1, knownLifts = ["squat", "bench", "deadlift"]) {
  const cache = new WorksheetCache(worksheet);
  const result = {
    duration: 0,
    weekly_schedule: createWeeklyArray(7, ""),
    training_days: [],
    exercises: [],
    exercisePositionMap: new Map(),
  };

  // Count weeks
  let weekCount = 0;
  let currentCol = 1;
  const maxColumns = Math.min(cache.worksheet.columnCount || 500, 500); // Support up to 500 columns
  while (currentCol < maxColumns) {
    const weekHeader = findWeekHeader(cache, currentCol);
    if (weekHeader) {
      weekCount++;
      currentCol +=
        TABLE_CONSTANTS.COLS_PER_DAY + TABLE_CONSTANTS.SPACING_BETWEEN_WEEKS;
    } else {
      break;
    }
  }
  result.duration = weekCount || 1;

  // Get all boundaries at once
  const allDayBoundaries = findDayBoundaries(cache, 1);
  const week1Days = allDayBoundaries.filter((d) => d.weekNum === 1);

  // Store week boundaries for reuse
  const weekBoundariesCache = new Map();
  weekBoundariesCache.set(0, week1Days);

  // Process week 1 structure
  week1Days.forEach((dayBoundary, dayIndex) => {
    const nextDay = week1Days[dayIndex + 1];
    const dayEndRow = nextDay ? nextDay.row - 1 : dayBoundary.row + 200;

    const trainingDay = {
      _id: generateContextualId(dayBoundary.dayName),
      name: dayBoundary.dayName,
      exercises: [],
    };

    const exerciseBoundaries = findExerciseBoundaries(
      cache,
      dayBoundary.row + 3,
      dayEndRow,
      1
    );

    exerciseBoundaries.forEach((exerciseBoundary, exerciseIndex) => {
      const nextExercise = exerciseBoundaries[exerciseIndex + 1];
      const exerciseEndRow = nextExercise ? nextExercise.row - 1 : dayEndRow;

      const parsed = parseExerciseName(exerciseBoundary.headerText, knownLifts);
      const exerciseData = parseExerciseDataWithBoundariesFast(
        cache,
        exerciseBoundary.row,
        exerciseEndRow,
        1,
        parsed,
        weekCount
      );

      exerciseData._id = generateContextualId(parsed.name);
      result.exercises.push(exerciseData);
      trainingDay.exercises.push(exerciseData._id);

      result.exercisePositionMap.set(exerciseData._id, {
        dayIndex,
        exerciseIndex,
        dayName: dayBoundary.dayName,
      });
    });

    result.training_days.push(trainingDay);

    const dayOfWeek = dayBoundary.dayNum - 1;
    if (dayOfWeek >= 0 && dayOfWeek < 7) {
      result.weekly_schedule[dayOfWeek] = trainingDay._id;
    }
  });

  // Parse weekly schedule
  const parsedSchedule = findAndParseWeeklySchedule(cache);
  const hasValidSchedule = parsedSchedule.some((day) => day !== "");

  if (hasValidSchedule) {
    result.weekly_schedule = createWeeklyArray(7, "");
    const trainingDaysByOrder = week1Days
      .map((dayBoundary) =>
        result.training_days.find((td) => td.name === dayBoundary.dayName)
      )
      .filter(Boolean);

    let trainingDayIndex = 0;
    for (let i = 0; i < 7; i++) {
      if (
        parsedSchedule[i] &&
        parsedSchedule[i].trim() !== "" &&
        trainingDayIndex < trainingDaysByOrder.length
      ) {
        result.weekly_schedule[i] = trainingDaysByOrder[trainingDayIndex]._id;
        trainingDayIndex++;
      }
    }
  }

  // Update weekly parameters from subsequent weeks
  for (let weekIdx = 1; weekIdx < weekCount; weekIdx++) {
    const weekCol =
      1 +
      weekIdx *
      (TABLE_CONSTANTS.COLS_PER_DAY + TABLE_CONSTANTS.SPACING_BETWEEN_WEEKS);

    // Cache week boundaries
    if (!weekBoundariesCache.has(weekIdx)) {
      const weekDays = findDayBoundaries(cache, weekCol).filter(
        (d) => d.weekNum === weekIdx + 1
      );
      weekBoundariesCache.set(weekIdx, weekDays);
    }

    const weekDays = weekBoundariesCache.get(weekIdx);

    weekDays.forEach((dayBoundary) => {
      const matchingDayIndex = week1Days.findIndex(
        (d) => d.dayName === dayBoundary.dayName
      );
      if (matchingDayIndex === -1) return;

      const nextDay = weekDays.find((d) => d.row > dayBoundary.row);
      const dayEndRow = nextDay ? nextDay.row - 1 : dayBoundary.row + 200;

      const exerciseBoundaries = findExerciseBoundaries(
        cache,
        dayBoundary.row + 3,
        dayEndRow,
        weekCol
      );

      exerciseBoundaries.forEach((exerciseBoundary, exerciseIndex) => {
        const matchingExercise = result.exercises.find((ex) => {
          const position = result.exercisePositionMap.get(ex._id);
          return (
            position &&
            position.dayIndex === matchingDayIndex &&
            position.exerciseIndex === exerciseIndex
          );
        });

        if (matchingExercise) {
          updateExerciseWeeklyDataFast(
            cache,
            exerciseBoundary.row,
            weekCol,
            weekIdx,
            matchingExercise,
            knownLifts
          );
        }
      });
    });
  }

  return result;
}

/**
 * Parse week execution data optimized
 */
function parseWeekExecutionData(
  cache,
  exerciseRow,
  weekCol,
  exercise,
  weekIdx
) {
  const { EXEC_WEIGHT, EXEC_REPS, EXEC_RPE, EXEC_NOTES } =
    TABLE_CONSTANTS.DAY_COLS;
  const weekData = {
    sets: [],
    note: "",
  };
  let currentRow = exerciseRow + 1;
  let totalSets = 0;
  // Calculate total sets for this specific week
  for (const sg of exercise.set_groups) {
    totalSets += sg.weekly_num_sets[weekIdx] || 0;
  }
  // Batch read all set data
  const sets = [];
  for (let setIdx = 0; setIdx < totalSets; setIdx++) {
    const rowData = cache.getRow(currentRow + setIdx);
    const execWeight = rowData[weekCol + EXEC_WEIGHT];
    const execReps = rowData[weekCol + EXEC_REPS];
    const execRpe = rowData[weekCol + EXEC_RPE];

    // Use explicit null/undefined checks to allow 0 values
    const hasWeight = execWeight !== null && execWeight !== undefined && execWeight !== "";
    const hasReps = execReps !== null && execReps !== undefined && execReps !== "";
    const hasRpe = execRpe !== null && execRpe !== undefined && execRpe !== "";

    sets.push({
      weight: hasWeight
        ? parseFloat(execWeight)
        : PARSE_CONFIG.DEFAULT_FILL_VALUE,
      reps: hasReps
        ? parseFloat(execReps)
        : PARSE_CONFIG.DEFAULT_FILL_VALUE,
      rpe: hasRpe
        ? parseFloat(execRpe)
        : PARSE_CONFIG.DEFAULT_FILL_VALUE,
      completed: hasWeight || hasReps || hasRpe,
    });
  }
  weekData.sets = sets;
  // Get notes from first set
  const notesValue = cache.getCellValue(currentRow, weekCol + EXEC_NOTES);
  if (notesValue) {
    weekData.note = notesValue.toString();
  }
  return weekData;
}

/**
 * Parse exercise recordings optimized
 */
function parseExerciseRecordings(
  worksheet,
  exercise,
  block,
  exercisePositionMap
) {
  const cache = new WorksheetCache(worksheet);
  const recordings = {
    exercise_id: exercise._id,
    weekly: Array(block.duration)
      .fill(null)
      .map(() => ({ sets: [], note: "" })),
  };

  const position = exercisePositionMap.get(exercise._id);
  if (!position) {
    warnLog(`No position info for exercise ${exercise.exercise_name}`);
    return recordings;
  }

  const { dayIndex, exerciseIndex } = position;

  // Cache for week boundaries
  const weekBoundariesCache = new Map();

  for (let weekIdx = 0; weekIdx < block.duration; weekIdx++) {
    const weekCol =
      1 +
      weekIdx *
      (TABLE_CONSTANTS.COLS_PER_DAY + TABLE_CONSTANTS.SPACING_BETWEEN_WEEKS);

    // Get or compute week boundaries
    if (!weekBoundariesCache.has(weekIdx)) {
      const dayBoundaries = findDayBoundaries(cache, weekCol);
      const weekDays = dayBoundaries.filter((d) => d.weekNum === weekIdx + 1);
      weekBoundariesCache.set(weekIdx, weekDays);
    }

    const weekDays = weekBoundariesCache.get(weekIdx);

    if (dayIndex < weekDays.length) {
      const targetDay = weekDays[dayIndex];
      const nextDay = weekDays[dayIndex + 1];
      const dayEndRow = nextDay ? nextDay.row - 1 : targetDay.row + 200;

      const exerciseBoundaries = findExerciseBoundaries(
        cache,
        targetDay.row + 3,
        dayEndRow,
        weekCol
      );

      if (exerciseIndex < exerciseBoundaries.length) {
        const targetExercise = exerciseBoundaries[exerciseIndex];
        const weekData = parseWeekExecutionData(
          cache,
          targetExercise.row,
          weekCol,
          exercise,
          weekIdx
        );
        recordings.weekly[weekIdx] = weekData;
      }
    }
  }

  return recordings;
}

/**
 * Parse a SigmaLifting XLSX workbook from a Buffer or Uint8Array.
 */
export const parseXlsxBufferToBundle = async (inputBuffer) => {
  try {
    debugLog("Starting parseXlsxBufferToBundle...");

    if (!inputBuffer || inputBuffer.byteLength === 0) {
      throw new Error("XLSX buffer is empty");
    }

    if (inputBuffer.byteLength > PARSE_CONFIG.MAX_FILE_SIZE) {
      throw new Error(
        `XLSX size exceeds limit of ${Math.round(
          PARSE_CONFIG.MAX_FILE_SIZE / 1024
        )}KB. File size: ${Math.round(
          inputBuffer.byteLength / 1024
        )}KB. Please try sharing via JSON.`
      );
    }
    debugLog("Creating new ExcelJS workbook...");

    const workbook = wrapError(
      () => new ExcelJS.Workbook(),
      "Failed to create workbook"
    );

    debugLog("Loading Excel file into workbook...");
    await wrapError(
      async () => await workbook.xlsx.load(inputBuffer),
      "Failed to load Excel file"
    );

    debugLog("Starting workbook sanitization...");
    sanitizeWorkbook(workbook);

    // Clear attribute cache for new parse
    attributeCache.clear();

    const result = {
      program: {
        _id: generateId(),
        name: "",
        description: "",
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString(),
        blocks: [],
        custom_anchored_lifts: [], // Initialize empty, will populate from 1RM profile
      },
      exercises: [],
      process: null,
      metadata: {
        exportedAt: new Date().toISOString(),
        version: "1.0",
        appName: "SigmaLifting",
      },
    };

    // Parse metadata and 1RM profile first
    const metadataSheet = workbook.getWorksheet("Metadata");
    const metadata = parseMetadata(metadataSheet);

    const oneRmProfileSheet = workbook.getWorksheet("1RM Profile");
    const oneRmProfile = parseOneRmProfile(oneRmProfileSheet);

    // Populate program's custom_anchored_lifts from parsed 1RM profile
    if (oneRmProfile.customLifts && Array.isArray(oneRmProfile.customLifts)) {
      result.program.custom_anchored_lifts = oneRmProfile.customLifts;
      debugLog(
        `[Excel Import] Program has ${oneRmProfile.customLifts.length} custom anchor lifts:`,
        oneRmProfile.customLifts
      );
    }

    debugLog("Finding block worksheets...");
    const blockSheets = [];

    // Use traditional for loop for better performance
    const worksheetCount = workbook.worksheets.length;
    for (let i = 0; i < worksheetCount; i++) {
      const ws = workbook.worksheets[i];
      if (
        ws?.name?.startsWith("Block ") &&
        !ws.name.includes("1RM") &&
        !ws.name.includes("RPE")
      ) {
        blockSheets.push(ws);
      }
    }

    debugLog(`Found ${blockSheets.length} block worksheets`);

    // Check for execution data early
    const hasExecutionData =
      metadata?.programOnly === true
        ? false
        : metadata?.programOnly === false
          ? true
          : checkForExecutionData(blockSheets);

    const blockDataArray = [];

    // Build known lifts array from 1RM Profile (lowercase for matching)
    // Sort by length (longest first) for correct matching of overlapping names
    const knownLifts = ["squat", "bench", "deadlift"];
    if (oneRmProfile.customLifts && Array.isArray(oneRmProfile.customLifts)) {
      knownLifts.push(...oneRmProfile.customLifts.map(l => l.toLowerCase()));
    }
    // Pre-sort once to avoid sorting on every parseExerciseName call
    knownLifts.sort((a, b) => b.length - a.length);

    // Process blocks
    for (let blockIndex = 0; blockIndex < blockSheets.length; blockIndex++) {
      const worksheet = blockSheets[blockIndex];
      const blockName = worksheet.name.replace(/^Block \d+ - /, "");

      const blockData = parseBlockWorksheet(
        worksheet,
        blockIndex,
        oneRmProfile.weekStartDay ?? 0,
        knownLifts
      );
      blockDataArray.push(blockData);

      const block = {
        _id: generateId(),
        name: blockName,
        program_id: result.program._id,
        duration: blockData.duration,
        weekly_schedule: blockData.weekly_schedule,
        training_days: blockData.training_days,
      };

      result.program.blocks.push(block);

      // Add exercises
      for (const exercise of blockData.exercises) {
        exercise.block_id = block._id;
        exercise.program_id = result.program._id;
        result.exercises.push(exercise);
      }

      // Clear attribute cache after each block to prevent memory issues
      // Blocks have different training patterns, so cache reuse between blocks is minimal
      attributeCache.clear();
    }

    // Set program name
    result.program.name =
      metadata && metadata.programName
        ? metadata.programName
        : `Imported Program - ${new Date().toLocaleDateString()}`;

    // Process execution data if present
    if (hasExecutionData) {
      result.process = {
        _id: generateId(),
        name:
          metadata && metadata.processName
            ? metadata.processName
            : `Imported Process - ${new Date().toLocaleDateString()}`,
        program_id: result.program._id,
        program_name: result.program.name,
        user_id: "",
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString(),
        start_date: new Date().toISOString().split("T")[0],
        config: {
          weight_unit:
            oneRmProfile.weightUnit || PARSE_CONFIG.DEFAULT_WEIGHT_UNIT,
          weight_rounding:
            oneRmProfile.weightRounding || PARSE_CONFIG.DEFAULT_WEIGHT_ROUNDING,
          week_start_day:
            oneRmProfile.weekStartDay ?? PARSE_CONFIG.DEFAULT_WEEK_START,
        },
        one_rm_profile: oneRmProfile.oneRmData,
        exercise_recordings: [],
      };

      // Handle blockly_one_rm
      if (result.process.one_rm_profile.enable_blockly_one_rm) {
        const numBlocks = result.program.blocks.length;
        const currentArray = result.process.one_rm_profile.blockly_one_rm || [];

        if (currentArray.length > numBlocks) {
          // Truncate: Excel has more blocks than program (e.g., trailing Block 3 data)
          // Keep only the first numBlocks entries to preserve valid data
          result.process.one_rm_profile.blockly_one_rm = currentArray.slice(0, numBlocks);
        } else if (currentArray.length < numBlocks) {
          // Extend: Excel has fewer blocks than program
          // Add default entries for missing blocks
          const missingCount = numBlocks - currentArray.length;
          const newEntries = Array(missingCount)
            .fill(null)
            .map(() => {
              const blockProfile = {
                squat: PARSE_CONFIG.DEFAULT_FILL_VALUE,
                bench: PARSE_CONFIG.DEFAULT_FILL_VALUE,
                deadlift: PARSE_CONFIG.DEFAULT_FILL_VALUE,
              };
              // Add custom anchor lifts to block-specific profiles
              // Use lowercase keys for consistency with how they're stored elsewhere
              if (oneRmProfile.customLifts && Array.isArray(oneRmProfile.customLifts)) {
                oneRmProfile.customLifts.forEach(liftName => {
                  blockProfile[liftName.toLowerCase()] = PARSE_CONFIG.DEFAULT_FILL_VALUE;
                });
              }
              return blockProfile;
            });
          result.process.one_rm_profile.blockly_one_rm = [...currentArray, ...newEntries];
        }
        // If lengths are equal, do nothing - data is correct
      }

      // Parse exercise recordings
      for (const exercise of result.exercises) {
        const blockIndex = result.program.blocks.findIndex(
          (b) => b._id === exercise.block_id
        );
        const worksheet = blockSheets[blockIndex];
        const recordings = parseExerciseRecordings(
          worksheet,
          exercise,
          result.program.blocks[blockIndex],
          blockDataArray[blockIndex].exercisePositionMap
        );

        if (recordings) {
          result.process.exercise_recordings.push(recordings);
        }
      }
    }

    return {
      success: true,
      data: result.process
        ? result
        : {
            program: result.program,
            exercises: result.exercises,
            metadata: result.metadata,
          },
      error: null,
    };
  } catch (error) {
    errorLog("Error parsing Excel file:", error);
    return {
      success: false,
      data: null,
      error: error.message || "Failed to parse Excel file",
    };
  }
};
