import ExcelJS from "exceljs";
import { generateExportFilename, truncateWorksheetName } from "./util.js";
import { RPE_CHART_DATA } from "../rpe.js";
import { calculatePrescriptionForSet } from "./prescriptionCalculator.js";
import { TABLE_CONSTANTS } from "./constants.js";
import {
  formatWeight,
  getEffectiveWeightRounding,
} from "./weightFormatting.js";

const CLI_APP_VERSION = "cli";

// Color definitions
const COLORS = {
  HEADER_GRAY: { argb: "FFD9D9D9" },
  WEEK_HEADER: { argb: "FFF2F2F2" },
  DAY_HEADER: { argb: "FFD9E2F3" },
  LIGHT_GRAY: { argb: "FFF5F5F5" },
  WHITE: { argb: "FFFFFFFF" },
  BLACK: { argb: "FF000000" },
};

// Performance optimization: Pre-create reusable style objects
const STYLES = {
  weekHeader: {
    font: { bold: true, size: 12 },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: { type: "pattern", pattern: "solid", fgColor: COLORS.WEEK_HEADER },
  },
  dayHeader: {
    font: { bold: true, size: 11 },
    alignment: { horizontal: "left", vertical: "middle" },
    fill: { type: "pattern", pattern: "solid", fgColor: COLORS.DAY_HEADER },
  },
  columnHeader: {
    font: { bold: true, size: 10 },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: { type: "pattern", pattern: "solid", fgColor: COLORS.HEADER_GRAY },
    border: {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    },
  },
  dataCell: {
    alignment: { horizontal: "center", vertical: "middle" },
    border: {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    },
  },
  notesCell: {
    alignment: { wrapText: true, vertical: "top", horizontal: "left" },
    font: { italic: true, size: 9 },
    border: {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    },
  },
};

// Performance: Create style IDs for faster comparison
const STYLE_IDS = {
  WEEK_HEADER: 'weekHeader',
  DAY_HEADER: 'dayHeader',
  COLUMN_HEADER: 'columnHeader',
  DATA_CELL: 'dataCell',
  NOTES_CELL: 'notesCell',
};

// Helper to get style ID from style object (avoids JSON stringify)
const getStyleId = (style) => {
  if (!style) return null;
  
  // Quick identification based on unique properties
  if (style.fill?.fgColor?.argb === "FFF2F2F2") return STYLE_IDS.WEEK_HEADER;
  if (style.fill?.fgColor?.argb === "FFD9E2F3") return STYLE_IDS.DAY_HEADER;
  if (style.fill?.fgColor?.argb === "FFD9D9D9") return STYLE_IDS.COLUMN_HEADER;
  if (style.alignment?.wrapText) return STYLE_IDS.NOTES_CELL;
  if (style.border) return STYLE_IDS.DATA_CELL;
  
  return 'custom';
};

// Helper to get set group attributes text
const getSetGroupAttributes = (setGroup, weekIndex, exercise, isDeloadWeek) => {
  const attributes = [];
  
  // If it's a deload week, only show deload information
  if (isDeloadWeek && exercise.deload_config?.enabled) {
    const deloadPercentage = exercise.deload_config.percentage || 85;
    const previousWeek = weekIndex > 0 ? weekIndex : 1;
    attributes.push(`deload ${deloadPercentage}% from week ${previousWeek}`);
    return attributes.join(". ");
  }
  
  // Check for backoff sets
  if (setGroup.backoff_config?.enabled) {
    const backoffType = setGroup.backoff_config.type;
    const dependsOnGroupId = setGroup.backoff_config.depends_on_set_group_id;
    let dependentSetNumber = "last set";
    
    if (dependsOnGroupId && exercise) {
      // Calculate which set number it depends on by finding the last set of the dependent group
      let setCounter = 0;
      for (const group of exercise.set_groups) {
        const numSets = group.weekly_num_sets?.[weekIndex] || 0;
        if (group.group_id === dependsOnGroupId) {
          dependentSetNumber = `set ${setCounter + numSets}`;
          break;
        }
        setCounter += numSets;
      }
    }
    
    if (backoffType === "%_of_weight") {
      const percentage = setGroup.backoff_config.weekly_percentage?.[weekIndex];
      if (percentage) {
        attributes.push(`backoff ${percentage}% from ${dependentSetNumber}`);
      }
    } else if (backoffType === "target_rpe") {
      const targetRpe = setGroup.backoff_config.weekly_rpe?.[weekIndex];
      if (targetRpe) {
        attributes.push(`backoff to RPE ${targetRpe} from ${dependentSetNumber}`);
      }
    }
  }
  
  // Check for fatigue drops
  if (setGroup.fatigue_drop_config?.enabled) {
    const rpeCap = setGroup.fatigue_drop_config.rpe_cap?.[weekIndex];
    const fatigueType = setGroup.fatigue_drop_config.type;
    
    if (rpeCap) {
      if (fatigueType === "%_of_weight") {
        const dropPercentage = setGroup.fatigue_drop_config.weekly_percentage?.[weekIndex];
        if (dropPercentage) {
          attributes.push(`Fatigue drop ${dropPercentage}% with RPE cap ${rpeCap}`);
        }
      } else if (fatigueType === "target_rpe") {
        const targetRpe = setGroup.fatigue_drop_config.weekly_rpe?.[weekIndex];
        if (targetRpe) {
          attributes.push(`Fatigue drop to RPE ${targetRpe} with RPE cap ${rpeCap}`);
        }
      }
    }
  }
  
  // Add weight percentage info for non-variable weight sets that aren't backoff or deload
  if (setGroup.variable_parameter !== "weight" && 
      !setGroup.backoff_config?.enabled && 
      !isDeloadWeek && 
      exercise.one_rm_anchor?.enabled) {
    
    // Check for mixed weight mode
    if (setGroup.mix_weight_config?.enabled) {
      const percentage = setGroup.mix_weight_config.weekly_weight_percentage?.[weekIndex];
      const absolute = setGroup.mix_weight_config.weekly_weight_absolute?.[weekIndex];
      const unit = setGroup.mix_weight_config.weight_unit || "kg";
      
      if (percentage !== undefined && absolute !== undefined) {
        attributes.push(`percentage: ${percentage}% absolute: ${absolute}${unit}`);
      }
    } 
    // Check for pure percentage mode
    else if (setGroup.weekly_weight_percentage?.[weekIndex] !== undefined) {
      const percentage = setGroup.weekly_weight_percentage[weekIndex];
      attributes.push(`percentage: ${percentage}%`);
    }
  }
  
  return attributes.join(". ");
};


// Helper function to reorder training_days based on weekly_schedule
const reorderTrainingDays = (block) => {
  const orderedDays = [];
  
  // Add training days in the order they appear in weekly_schedule
  block.weekly_schedule.forEach(dayId => {
    if (dayId && dayId !== "") {
      const day = block.training_days.find(td => td._id === dayId);
      if (day) {
        orderedDays.push(day);
      }
    }
  });
  
  // Return a new block object with reordered training_days
  return {
    ...block,
    training_days: orderedDays
  };
};

// OPTIMIZED: Batch data collection for entire block
const collectBlockData = (block, process, exerciseMap, blockIndex) => {
  const blockWithOrderedDays = reorderTrainingDays(block);
  const blockData = {
    rows: [],
    merges: [],
    columnWidths: new Map(),
  };
  
  let currentCol = 1;
  const hasTrainingDays = blockWithOrderedDays.training_days && blockWithOrderedDays.training_days.length > 0;
  
  if (hasTrainingDays) {
    // Process each week
    for (let weekIdx = 0; weekIdx < blockWithOrderedDays.duration; weekIdx++) {
      if (weekIdx > 0) {
        currentCol += TABLE_CONSTANTS.SPACING_BETWEEN_WEEKS;
      }
      
      const weekStartCol = currentCol;
      collectWeekData(blockData, blockWithOrderedDays, weekIdx, process, exerciseMap, weekStartCol, blockIndex);
      currentCol += TABLE_CONSTANTS.COLS_PER_DAY;
    }
    
    // Add spacing before weekly schedule
    currentCol += TABLE_CONSTANTS.SPACING_BETWEEN_WEEKS;
  }
  
  // Collect weekly schedule data
  collectWeeklyScheduleData(blockData, blockWithOrderedDays, process, currentCol);
  
  return blockData;
};

// OPTIMIZED: Apply collected data to worksheet in batch with style batching
const applyBlockDataToWorksheet = (worksheet, blockData) => {
  // Performance optimization: Batch operations by type to minimize proxy overhead
  
  // 1. Set all column widths at once
  blockData.columnWidths.forEach((width, col) => {
    worksheet.getColumn(col).width = width;
  });
  
  // 2. Prepare style groups for batch application
  const styleGroups = new Map(); // styleId -> [{row, col, style}, ...]
  const borderRegions = [];       // Collect border regions for batch application
  
  // 3. First pass: Set all values and collect style metadata
  const allValues = []; // Array of row values for batch setting
  
  blockData.rows.forEach((rowData, rowIndex) => {
    if (rowData && Object.keys(rowData).length > 0) {
      const actualRowNum = rowIndex + 1;
      const values = [];
      let maxCol = 0;
      
      // Find max column for this row
      Object.keys(rowData).forEach(col => {
        const colNum = parseInt(col);
        if (colNum > maxCol) maxCol = colNum;
      });
      
      // Build values array and collect style info
      for (let colNum = 1; colNum <= maxCol; colNum++) {
        const cellData = rowData[colNum];
        if (cellData) {
          values[colNum] = cellData.value;
          
          // Collect style metadata for batch processing
          if (cellData.style) {
            // Use style ID for efficient grouping (avoids JSON operations)
            const styleId = getStyleId(cellData.style);
            
            if (!styleGroups.has(styleId)) {
              styleGroups.set(styleId, []);
            }
            styleGroups.get(styleId).push({ 
              row: actualRowNum, 
              col: colNum,
              style: cellData.style 
            });
            
            // Collect border info separately (borders are expensive)
            if (cellData.style.border) {
              borderRegions.push({
                row: actualRowNum,
                col: colNum,
                border: cellData.style.border
              });
            }
          }
        }
      }
      
      // Store row values for batch setting
      allValues[actualRowNum] = values;
    }
  });
  
  // 4. Batch set all cell values at once
  allValues.forEach((values, rowNum) => {
    if (values && values.length > 0) {
      const row = worksheet.getRow(rowNum);
      row.values = values;
    }
  });
  
  // 5. Apply styles in batches by style type
  // Process each style group efficiently
  styleGroups.forEach((cells, styleId) => {
    if (cells.length === 0) return;
    
    // Get the appropriate style object
    let baseStyle;
    if (styleId === 'custom') {
      // For custom styles, use the first cell's style
      baseStyle = cells[0]?.style;
    } else {
      // Use predefined styles for known patterns
      baseStyle = STYLES[styleId];
    }
    
    if (!baseStyle) return;
    
    // Apply styles in batch to minimize proxy overhead
    // Group consecutive cells when possible
    cells.forEach(({ row, col, style }) => {
      const cell = worksheet.getRow(row).getCell(col);
      
      // For predefined styles, apply the complete style object
      if (styleId !== 'custom') {
        // Apply only non-border properties first
        if (baseStyle.font) cell.font = baseStyle.font;
        if (baseStyle.alignment) cell.alignment = baseStyle.alignment;
        if (baseStyle.fill) cell.fill = baseStyle.fill;
      } else {
        // For custom styles, apply the actual style
        if (style.font) cell.font = style.font;
        if (style.alignment) cell.alignment = style.alignment;
        if (style.fill) cell.fill = style.fill;
      }
    });
  });
  
  // 6. Apply borders in batch at the end (most expensive operation)
  // Group adjacent cells with same border style
  const borderGroups = new Map(); // borderKey -> ranges[]
  
  borderRegions.forEach(({ row, col, border }) => {
    const borderKey = JSON.stringify(border);
    if (!borderGroups.has(borderKey)) {
      borderGroups.set(borderKey, []);
    }
    borderGroups.get(borderKey).push({ row, col });
  });
  
  // Apply borders by group with range detection
  borderGroups.forEach((cells, borderKey) => {
    const border = JSON.parse(borderKey);
    
    // Sort cells to find contiguous ranges
    cells.sort((a, b) => a.row - b.row || a.col - b.col);
    
    // Detect contiguous horizontal ranges for more efficient application
    const ranges = [];
    let currentRange = null;
    
    cells.forEach(({ row, col }) => {
      if (!currentRange || 
          currentRange.row !== row || 
          col !== currentRange.endCol + 1) {
        // Start new range
        currentRange = { row, startCol: col, endCol: col };
        ranges.push(currentRange);
      } else {
        // Extend current range
        currentRange.endCol = col;
      }
    });
    
    // Apply borders to ranges (reduces cell access)
    ranges.forEach(range => {
      if (range.startCol === range.endCol) {
        // Single cell
        const cell = worksheet.getRow(range.row).getCell(range.startCol);
        cell.border = border;
      } else {
        // Range of cells - apply to each (ExcelJS doesn't have range border API)
        for (let col = range.startCol; col <= range.endCol; col++) {
          const cell = worksheet.getRow(range.row).getCell(col);
          cell.border = border;
        }
      }
    });
  });
  
  // 8. Apply all merges at once (already optimized)
  blockData.merges.forEach(merge => {
    worksheet.mergeCells(merge.top, merge.left, merge.bottom, merge.right);
  });
};

// Function to populate a single block worksheet with self-contained days
export const populateBlockWorksheet = (
  worksheet,
  block,
  blockIndex,
  program,
  process,
  exerciseMap
) => {
  // Collect all data first
  const blockData = collectBlockData(block, process, exerciseMap, blockIndex);
  
  // Apply all data in batch
  applyBlockDataToWorksheet(worksheet, blockData);
};

// OPTIMIZED: Collect weekly schedule data
const collectWeeklyScheduleData = (blockData, block, process, startCol) => {
  let currentRow = 1;
  
  // Set column widths
  blockData.columnWidths.set(startCol, 15);
  blockData.columnWidths.set(startCol + 1, 20);
  
  // Add header with -> prefix for easy parsing
  if (!blockData.rows[currentRow - 1]) blockData.rows[currentRow - 1] = [];
  blockData.rows[currentRow - 1][startCol] = {
    value: "-> Weekly Schedule",
    style: { 
      font: { bold: true, size: 14 },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: { type: "pattern", pattern: "solid", fgColor: COLORS.WEEK_HEADER },
    },
  };
  blockData.merges.push({
    top: currentRow,
    left: startCol,
    bottom: currentRow,
    right: startCol + 1,
  });
  currentRow++;
  
  // Skip the "Day 1 - Day 7" subheader as requested
  // currentRow++ is not needed since we're not adding this row
  
  // Get day names based on week_start_day
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const weekStartDay = process.config.week_start_day ?? 0; // 0 = Sunday by default
  
  // Reorder day names based on week_start_day
  const orderedDayNames = [];
  for (let i = 0; i < 7; i++) {
    const dayIndex = (weekStartDay + i) % 7;
    orderedDayNames.push(dayNames[dayIndex]);
  }
  
  // Skip the Day and Training Day headers as requested
  // currentRow++ is not needed since we're not adding this row
  
  // Add each day
  orderedDayNames.forEach((dayName, index) => {
    if (!blockData.rows[currentRow - 1]) blockData.rows[currentRow - 1] = [];
    
    blockData.rows[currentRow - 1][startCol] = {
      value: dayName,
      style: {
        border: {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        },
      },
    };
    
    const trainingDayId = block.weekly_schedule[index];
    const trainingDay = block.training_days.find(td => td._id === trainingDayId);
    
    let trainingDayValue = "";
    if (trainingDay) {
      const actualDayNumber = block.training_days.findIndex(td => td._id === trainingDay._id) + 1;
      trainingDayValue = `Day ${actualDayNumber} - ${trainingDay.name}`;
    }
    
    blockData.rows[currentRow - 1][startCol + 1] = {
      value: trainingDayValue,
      style: {
        border: {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        },
      },
    };
    
    currentRow++;
  });
};

// OPTIMIZED: Collect week data in batch
const collectWeekData = (blockData, block, weekIdx, process, exerciseMap, startCol, blockIndex) => {
  let currentRow = 1;
  
  // Set column widths once for this week
  const dayColumns = TABLE_CONSTANTS.DAY_COLS;
  blockData.columnWidths.set(startCol + dayColumns.SET_NO, TABLE_CONSTANTS.SET_NO_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.PRESC_WEIGHT, TABLE_CONSTANTS.WEIGHT_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.PRESC_REPS, TABLE_CONSTANTS.REPS_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.PRESC_RPE, TABLE_CONSTANTS.RPE_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.PRESC_ATTRIBUTES, TABLE_CONSTANTS.ATTRIBUTES_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.PRESC_NOTES, TABLE_CONSTANTS.NOTES_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.EXEC_WEIGHT, TABLE_CONSTANTS.WEIGHT_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.EXEC_REPS, TABLE_CONSTANTS.REPS_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.EXEC_RPE, TABLE_CONSTANTS.RPE_COL_WIDTH);
  blockData.columnWidths.set(startCol + dayColumns.EXEC_NOTES, TABLE_CONSTANTS.NOTES_COL_WIDTH);
  
  // Process each training day
  block.training_days.forEach((day, dayIndex) => {
    if (dayIndex > 0) {
      currentRow += TABLE_CONSTANTS.SPACING_BETWEEN_DAYS;
    }
    
    const dayStartRow = currentRow;
    currentRow = collectDayData(blockData, day, dayIndex, weekIdx, process, exerciseMap, dayStartRow, startCol, block.duration, blockIndex);
  });
};

// OPTIMIZED: Collect day data in batch
const collectDayData = (blockData, day, dayIndex, weekIdx, process, exerciseMap, startRow, startCol, totalWeeksInBlock, blockIndex) => {
  let currentRow = startRow;
  const weightUnit = process.config.weight_unit?.toLowerCase() || "lbs";
  const weightRounding = getEffectiveWeightRounding(process.config);
  
  // Initialize row if it doesn't exist
  if (!blockData.rows[currentRow - 1]) {
    blockData.rows[currentRow - 1] = [];
  }
  
  // Week/Day header
  blockData.rows[currentRow - 1][startCol] = {
    value: `$ Week ${weekIdx + 1} - Day ${dayIndex + 1} - ${day.name}`,
    style: STYLES.weekHeader,
  };
  blockData.merges.push({
    top: currentRow,
    left: startCol,
    bottom: currentRow,
    right: startCol + TABLE_CONSTANTS.COLS_PER_DAY - 1,
  });
  currentRow++;
  
  // Prescription/Execution headers
  const dayColumns = TABLE_CONSTANTS.DAY_COLS;
  if (!blockData.rows[currentRow - 1]) blockData.rows[currentRow - 1] = [];
  blockData.rows[currentRow - 1][startCol + dayColumns.PRESC_WEIGHT] = {
    value: "Prescription",
    style: { font: { bold: true }, alignment: { horizontal: "center", vertical: "middle" } },
  };
  blockData.rows[currentRow - 1][startCol + dayColumns.EXEC_WEIGHT] = {
    value: "Execution",
    style: { font: { bold: true }, alignment: { horizontal: "center", vertical: "middle" } },
  };
  blockData.merges.push(
    {
      top: currentRow,
      left: startCol + dayColumns.PRESC_WEIGHT,
      bottom: currentRow,
      right: startCol + dayColumns.PRESC_NOTES,
    },
    {
      top: currentRow,
      left: startCol + dayColumns.EXEC_WEIGHT,
      bottom: currentRow,
      right: startCol + dayColumns.EXEC_NOTES,
    }
  );
  currentRow++;
  
  // Column headers
  if (!blockData.rows[currentRow - 1]) blockData.rows[currentRow - 1] = [];
  const headers = ["set no.", `weight(${weightUnit})`, "reps", "rpe", "attributes", "notes", `weight(${weightUnit})`, "reps", "rpe", "notes"];
  headers.forEach((header, idx) => {
    blockData.rows[currentRow - 1][startCol + idx] = {
      value: header,
      style: STYLES.columnHeader,
    };
  });
  currentRow++;
  
  // Process exercises for this day
  day.exercises.forEach((exerciseId, exerciseIndex) => {
    if (exerciseIndex > 0) {
      currentRow++; // Add spacing between exercises
    }
    
    const exercise = exerciseMap.get(exerciseId);
    if (!exercise) return;
    
    currentRow = collectExerciseData(
      blockData,
      exercise,
      weekIdx,
      process,
      startCol,
      currentRow,
      totalWeeksInBlock,
      blockIndex,
      weightRounding
    );
  });
  
  return currentRow; // Return the final row number used
};

// OPTIMIZED: Collect exercise data in batch
const collectExerciseData = (
  blockData,
  exercise,
  weekIdx,
  process,
  startCol,
  startRow,
  totalWeeksInBlock,
  blockIndex,
  weightRounding
) => {
  let currentRow = startRow;
  
  // Check if this is a deload week
  const isDeloadWeek = exercise.deload_config?.enabled && 
                       totalWeeksInBlock && 
                       weekIdx === totalWeeksInBlock - 1;
  
  // Find exercise recording
  const exerciseRecording = process.exercise_recordings.find(
    (rec) => rec.exercise_id === exercise._id
  );
  const weekData = exerciseRecording?.weekly?.[weekIdx];
  
  // Build exercise name with metadata
  let exerciseName = `# ${exercise.exercise_name}`;
  
  // Build the info parts for the parentheses
  const infoParts = [];
  
  // Add 1RM anchor information if enabled
  if (exercise.one_rm_anchor?.enabled) {
    const ratio = exercise.one_rm_anchor.ratio || 1.0;
    // Ensure lift_type is always lowercase, handling existing capitalized data
    const liftType = (exercise.one_rm_anchor.lift_type || "squat").toLowerCase();
    const percentage = (ratio * 100).toFixed(0);
    infoParts.push(`max: ${percentage}% ${liftType}`);
  }
  
  // Check for deload configuration
  if (exercise.deload_config?.enabled && exercise.deload_config.percentage) {
    infoParts.push(`deload week: ${exercise.deload_config.percentage}%`);
  }
  
  // Add set groups information
  // CRITICAL: Keep all set counts including zeros to preserve structural information
  // The number of elements (split by '+') indicates the number of set groups
  // Week 1 metadata is the source of truth for exercise structure during import
  const setGroupsInfo = exercise.set_groups
    .map(setGroup => setGroup.weekly_num_sets?.[weekIdx] || 0)
    .join('+');

  // Always add set groups info, even "0" or "0+0", to preserve structure
  infoParts.push(`set groups: ${setGroupsInfo}`);
  
  // Combine all info parts
  if (infoParts.length > 0) {
    exerciseName += ` (${infoParts.join(', ')})`;
  }
  
  // Add exercise header
  if (!blockData.rows[currentRow - 1]) blockData.rows[currentRow - 1] = [];
  blockData.rows[currentRow - 1][startCol] = {
    value: exerciseName,
    style: STYLES.dayHeader,
  };
  blockData.merges.push({
    top: currentRow,
    left: startCol,
    bottom: currentRow,
    right: startCol + TABLE_CONSTANTS.COLS_PER_DAY - 1,
  });
  currentRow++;
  
  // Calculate and populate sets
  const allSets = [];
  let setIndex = 0;
  
  exercise.set_groups.forEach((setGroup, groupIdx) => {
    const numSets = setGroup.weekly_num_sets?.[weekIdx] || 0;
    
    for (let i = 0; i < numSets; i++) {
      const prescription = calculatePrescriptionForSet(
        exercise,
        weekIdx,
        setIndex,
        process.one_rm_profile,
        blockIndex,
        weightRounding,
        allSets,
        process.config.weight_unit || "lbs",
        totalWeeksInBlock,
        exerciseRecording
      );
      
      const executedSet = weekData?.sets?.[setIndex];

      // Store for backoff calculations
      // Only use execution data if the set was actually completed (not just default -1 values)
      if (executedSet && executedSet.completed) {
        allSets.push({
          ...executedSet,
          prescribedWeight: prescription.weight,
          prescribedReps: prescription.reps,
          prescribedRpe: prescription.rpe
        });
      } else {
        allSets.push({
          weight: prescription.weight,
          reps: prescription.reps,
          rpe: prescription.rpe,
          completed: false,
          prescribedWeight: prescription.weight,
          prescribedReps: prescription.reps,
          prescribedRpe: prescription.rpe
        });
      }
      
      // Collect set data
      collectSetData(blockData, setGroup, prescription, executedSet, weekData, 
                    startCol, currentRow, setIndex, i, weekIdx, exercise, isDeloadWeek);
      
      currentRow++;
      setIndex++;
    }
  });
  
  return currentRow;
};

// OPTIMIZED: Collect set data
const collectSetData = (blockData, setGroup, prescription, executedSet, weekData, 
                        startCol, currentRow, setIndex, setGroupIndex, weekIdx, exercise, isDeloadWeek) => {
  const dayColumns = TABLE_CONSTANTS.DAY_COLS;
  
  if (!blockData.rows[currentRow - 1]) blockData.rows[currentRow - 1] = [];
  
  const rowData = blockData.rows[currentRow - 1];
  
  // Set all cells for this row
  rowData[startCol + dayColumns.SET_NO] = {
    value: setIndex + 1,
    style: STYLES.dataCell,
  };
  
  rowData[startCol + dayColumns.PRESC_WEIGHT] = {
    value: formatWeight(prescription.weight),
    style: STYLES.dataCell,
  };
  
  rowData[startCol + dayColumns.PRESC_REPS] = {
    value: prescription.reps ?? "",
    style: STYLES.dataCell,
  };
  
  rowData[startCol + dayColumns.PRESC_RPE] = {
    value: prescription.rpe ?? "",
    style: STYLES.dataCell,
  };
  
  // Attributes
  const attributes = getSetGroupAttributes(setGroup, weekIdx, exercise, isDeloadWeek);
  if (attributes) {
    rowData[startCol + dayColumns.PRESC_ATTRIBUTES] = {
      value: attributes,
      style: {
        ...STYLES.dataCell,
        alignment: { wrapText: true, vertical: "middle", horizontal: "left" },
        font: { size: 9 },
      },
    };
  } else {
    rowData[startCol + dayColumns.PRESC_ATTRIBUTES] = {
      value: "",
      style: STYLES.dataCell,
    };
  }
  
  // Program notes
  if (setGroupIndex === 0) {
    const programNote = setGroup.weekly_notes?.[weekIdx];
    if (programNote && programNote.trim()) {
      rowData[startCol + dayColumns.PRESC_NOTES] = {
        value: programNote,
        style: STYLES.notesCell,
      };
    } else {
      rowData[startCol + dayColumns.PRESC_NOTES] = {
        value: "",
        style: STYLES.dataCell,
      };
    }
  } else {
    rowData[startCol + dayColumns.PRESC_NOTES] = {
      value: "",
      style: STYLES.dataCell,
    };
  }
  
  // Execution data
  if (executedSet?.completed) {
    rowData[startCol + dayColumns.EXEC_WEIGHT] = {
      value: formatWeight(executedSet.weight),
      style: STYLES.dataCell,
    };
    rowData[startCol + dayColumns.EXEC_REPS] = {
      value: executedSet.reps ?? "",
      style: STYLES.dataCell,
    };
    rowData[startCol + dayColumns.EXEC_RPE] = {
      value: executedSet.rpe ?? "",
      style: STYLES.dataCell,
    };
  } else {
    rowData[startCol + dayColumns.EXEC_WEIGHT] = { value: "", style: STYLES.dataCell };
    rowData[startCol + dayColumns.EXEC_REPS] = { value: "", style: STYLES.dataCell };
    rowData[startCol + dayColumns.EXEC_RPE] = { value: "", style: STYLES.dataCell };
  }
  
  // Process notes
  if (setIndex === 0 && weekData?.note && weekData.note.trim()) {
    rowData[startCol + dayColumns.EXEC_NOTES] = {
      value: weekData.note,
      style: STYLES.notesCell,
    };
  } else {
    rowData[startCol + dayColumns.EXEC_NOTES] = {
      value: "",
      style: STYLES.dataCell,
    };
  }
};

// Function to create 1RM Profile worksheet
export const create1RMProfileWorksheet = (worksheet, process, program) => {
  const weightRounding = getEffectiveWeightRounding(process.config);

  // Set column widths
  worksheet.getColumn(1).width = 15;
  worksheet.getColumn(2).width = 12;
  worksheet.getColumn(3).width = 12;
  worksheet.getColumn(4).width = 12;
  worksheet.getColumn(5).width = 20;

  // Title
  const titleCell = worksheet.getCell(1, 1);
  titleCell.value = "1RM Profile";
  titleCell.font = { bold: true, size: 16 };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  worksheet.mergeCells(1, 1, 1, 5);

  let currentRow = 3;

  // Configuration section
  const configHeaderCell = worksheet.getCell(currentRow, 1);
  configHeaderCell.value = "Configuration";
  configHeaderCell.font = { bold: true, size: 14 };
  configHeaderCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: COLORS.WEEK_HEADER,
  };
  worksheet.mergeCells(currentRow, 1, currentRow, 5);
  currentRow++;

  // Weight unit
  const unitCell = worksheet.getCell(currentRow, 1);
  unitCell.value = "Weight Unit:";
  unitCell.font = { bold: true };
  
  const unitValueCell = worksheet.getCell(currentRow, 2);
  unitValueCell.value = process.config.weight_unit?.toUpperCase() || "LBS";
  unitValueCell.alignment = { horizontal: "center", vertical: "middle" };
  currentRow++;

  // Weight rounding
  const roundingCell = worksheet.getCell(currentRow, 1);
  roundingCell.value = "Weight Rounding:";
  roundingCell.font = { bold: true };
  
  const roundingValueCell = worksheet.getCell(currentRow, 2);
  roundingValueCell.value = weightRounding;
  roundingValueCell.alignment = { horizontal: "center", vertical: "middle" };
  currentRow++;

  // Day 1 starts on
  const dayStartCell = worksheet.getCell(currentRow, 1);
  dayStartCell.value = "Day 1 starts on:";
  dayStartCell.font = { bold: true };
  
  const dayStartValueCell = worksheet.getCell(currentRow, 2);
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const weekStartDay = process.config.week_start_day ?? 0; // 0 = Sunday by default
  dayStartValueCell.value = dayNames[weekStartDay];
  dayStartValueCell.alignment = { horizontal: "center", vertical: "middle" };
  currentRow += 2; // Add spacing

  // Global 1RM Profile section
  const globalHeaderCell = worksheet.getCell(currentRow, 1);
  globalHeaderCell.value = "Global 1RM Profile";
  globalHeaderCell.font = { bold: true, size: 14 };
  globalHeaderCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: COLORS.DAY_HEADER,
  };
  worksheet.mergeCells(currentRow, 1, currentRow, 5);
  currentRow++;

  // Global headers
  const weightUnit = process.config.weight_unit?.toLowerCase() || "lbs";
  const globalHeaders = ["Exercise", `1RM (${weightUnit})`, "", "", ""];
  globalHeaders.forEach((header, idx) => {
    if (idx < 2) {
      const cell = worksheet.getCell(currentRow, idx + 1);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: COLORS.HEADER_GRAY,
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  });
  currentRow++;

  // Global 1RM data
  const globalOneRm = process.one_rm_profile;

  // Build dynamic lift arrays: reserved lifts + custom lifts from program
  const exercises = ["Squat", "Bench", "Deadlift"];
  const exerciseKeys = ["squat", "bench", "deadlift"];

  // Add custom anchored lifts from program if they exist
  if (program.custom_anchored_lifts && Array.isArray(program.custom_anchored_lifts)) {
    program.custom_anchored_lifts.forEach(liftName => {
      exercises.push(liftName); // Display name with exact capitalization
      exerciseKeys.push(liftName.toLowerCase()); // Key in one_rm_profile (MUST be lowercase)
    });
  }

  // Export all lifts (reserved + custom)
  exercises.forEach((exercise, idx) => {
    const exerciseKey = exerciseKeys[idx];
    const cell1 = worksheet.getCell(currentRow, 1);
    const cell2 = worksheet.getCell(currentRow, 2);

    cell1.value = exercise;
    cell1.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    const oneRmValue = globalOneRm[exerciseKey];
    cell2.value = (oneRmValue !== undefined && oneRmValue !== null && oneRmValue > 0)
      ? formatWeight(oneRmValue)
      : "";
    cell2.alignment = { horizontal: "center", vertical: "middle" };
    cell2.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    currentRow++;
  });

  currentRow += 2; // Add spacing

  // Block-specific 1RM section
  const blockSpecificHeaderCell = worksheet.getCell(currentRow, 1);
  blockSpecificHeaderCell.value = "Block-Specific 1RM";
  blockSpecificHeaderCell.font = { bold: true, size: 14 };
  blockSpecificHeaderCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: COLORS.DAY_HEADER,
  };
  worksheet.mergeCells(currentRow, 1, currentRow, 5);
  currentRow++;

  // Block-specific enabled status
  const enabledCell = worksheet.getCell(currentRow, 1);
  enabledCell.value = "Block-specific 1RM enabled:";
  enabledCell.font = { bold: true };
  
  const statusCell = worksheet.getCell(currentRow, 2);
  statusCell.value = globalOneRm.enable_blockly_one_rm ? "Yes" : "No";
  statusCell.font = { bold: true, color: { argb: globalOneRm.enable_blockly_one_rm ? "FF008000" : "FFFF0000" } };
  currentRow += 2;

  // If block-specific is enabled, show the block data
  if (globalOneRm.enable_blockly_one_rm && globalOneRm.blockly_one_rm && globalOneRm.blockly_one_rm.length > 0) {
    // Build dynamic headers: reserved lifts + custom lifts
    const blockHeaders = ["Block", `Squat (${weightUnit})`, `Bench (${weightUnit})`, `Deadlift (${weightUnit})`];

    // Add custom lift headers from program
    if (program.custom_anchored_lifts && Array.isArray(program.custom_anchored_lifts)) {
      program.custom_anchored_lifts.forEach(liftName => {
        blockHeaders.push(`${liftName} (${weightUnit})`);
      });
    }

    // Write all headers
    blockHeaders.forEach((header, idx) => {
      const cell = worksheet.getCell(currentRow, idx + 1);
      cell.value = header;
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: COLORS.HEADER_GRAY,
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
    currentRow++;

    // Build lift keys array once (same as global section)
    const blockLiftKeys = ["squat", "bench", "deadlift"];
    if (program.custom_anchored_lifts && Array.isArray(program.custom_anchored_lifts)) {
      blockLiftKeys.push(...program.custom_anchored_lifts.map(l => l.toLowerCase()));
    }

    // Block-specific 1RM data - export only blocks that exist in program
    // Trim blockly_one_rm to match actual program block count to avoid exporting phantom blocks
    const actualBlockCount = program.blocks.length;
    const blocksToExport = globalOneRm.blockly_one_rm.slice(0, actualBlockCount);

    blocksToExport.forEach((blockOneRm, blockIdx) => {
      // Write block number in first column
      const blockCell = worksheet.getCell(currentRow, 1);
      blockCell.value = `Block ${blockIdx + 1}`;
      blockCell.alignment = { horizontal: "center", vertical: "middle" };
      blockCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      // Write values for all lifts (columns 2+)
      blockLiftKeys.forEach((liftKey, idx) => {
        const cell = worksheet.getCell(currentRow, idx + 2);
        const liftValue = blockOneRm[liftKey];
        cell.value = (liftValue !== undefined && liftValue !== null && liftValue > 0)
          ? formatWeight(liftValue)
          : "";
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      currentRow++;
    });
  } else if (globalOneRm.enable_blockly_one_rm) {
    // Block-specific is enabled but no data available
    const noDataCell = worksheet.getCell(currentRow, 1);
    noDataCell.value = "No block-specific 1RM data available";
    noDataCell.font = { italic: true };
    noDataCell.alignment = { horizontal: "center", vertical: "middle" };
    worksheet.mergeCells(currentRow, 1, currentRow, 4);
  }
};

// Function to create RPE Chart worksheet
export const createRPEChartWorksheet = (worksheet) => {
  // Set column widths
  worksheet.getColumn(1).width = 12;
  for (let i = 2; i <= 10; i++) {
    worksheet.getColumn(i).width = 8;
  }

  // Title
  const titleCell = worksheet.getCell(1, 1);
  titleCell.value = "RPE Chart";
  titleCell.font = { bold: true, size: 16 };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  worksheet.mergeCells(1, 1, 1, 10);

  // Headers
  const headerRow = 3;
  worksheet.getCell(headerRow, 1).value = "RPE\\Reps";
  for (let reps = 1; reps <= 9; reps++) {
    worksheet.getCell(headerRow, reps + 1).value = reps;
  }

  // Style headers
  for (let col = 1; col <= 10; col++) {
    const cell = worksheet.getCell(headerRow, col);
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: COLORS.DAY_HEADER,
    };
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  }

  // Add RPE data
  const rpeValues = [10, 9.5, 9, 8.5, 8, 7.5, 7, 6.5, 6];
  rpeValues.forEach((rpe, idx) => {
    const row = headerRow + idx + 1;
    worksheet.getCell(row, 1).value = rpe;
    worksheet.getCell(row, 1).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    for (let reps = 1; reps <= 9; reps++) {
      const cell = worksheet.getCell(row, reps + 1);
      const rpeKey = rpe.toString();
      const repsKey = reps.toString();
      const value = RPE_CHART_DATA[rpeKey]?.[repsKey];
      cell.value = value ? (value * 100).toFixed(1) : "";
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  });
};

// Function to create Metadata worksheet
const createMetadataWorksheet = (worksheet, process, program) => {
  // Set column widths
  worksheet.getColumn(1).width = 20;
  worksheet.getColumn(2).width = 50;

  // Title
  const titleCell = worksheet.getCell(1, 1);
  titleCell.value = "Metadata";
  titleCell.font = { bold: true, size: 16 };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  worksheet.mergeCells(1, 1, 1, 2);

  let currentRow = 3;

  // Add metadata fields
  const metadataFields = [
    { label: "Program Name:", value: program.name },
    { label: "Process Name:", value: process.name },
    { label: "Program Only:", value: "No" }, // No for process exports (includes execution data)
    { label: "Export Date:", value: new Date().toISOString() },
    { label: "App Version:", value: CLI_APP_VERSION },
  ];

  metadataFields.forEach(field => {
    const labelCell = worksheet.getCell(currentRow, 1);
    labelCell.value = field.label;
    labelCell.font = { bold: true };
    labelCell.alignment = { horizontal: "left", vertical: "middle" };
    labelCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    const valueCell = worksheet.getCell(currentRow, 2);
    valueCell.value = field.value;
    valueCell.alignment = { horizontal: "left", vertical: "middle" };
    valueCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    currentRow++;
  });

  // Add a note about the programOnly field
  currentRow += 2;
  const noteCell = worksheet.getCell(currentRow, 1);
  noteCell.value = "Note: 'Program Only' indicates whether this export contains only program data (Yes) or includes process/execution data (No)";
  noteCell.font = { italic: true, size: 10 };
  noteCell.alignment = { wrapText: true, vertical: "top" };
  worksheet.mergeCells(currentRow, 1, currentRow, 2);
};

// TABLE_CONSTANTS are now imported from ./constants.js to avoid circular dependencies

const createProgramMetadataWorksheet = (worksheet, program) => {
  worksheet.getColumn(1).width = 20;
  worksheet.getColumn(2).width = 50;

  const titleCell = worksheet.getCell(1, 1);
  titleCell.value = "Metadata";
  titleCell.font = { bold: true, size: 16 };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  worksheet.mergeCells(1, 1, 1, 2);

  let currentRow = 3;

  const metadataFields = [
    { label: "Program Name:", value: program.name },
    { label: "Process Name:", value: "" },
    { label: "Program Only:", value: "Yes" },
    { label: "Export Date:", value: new Date().toISOString() },
    { label: "App Version:", value: CLI_APP_VERSION },
  ];

  metadataFields.forEach((field) => {
    const labelCell = worksheet.getCell(currentRow, 1);
    labelCell.value = field.label;
    labelCell.font = { bold: true };
    labelCell.alignment = { horizontal: "left", vertical: "middle" };
    labelCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    const valueCell = worksheet.getCell(currentRow, 2);
    valueCell.value = field.value;
    valueCell.alignment = { horizontal: "left", vertical: "middle" };
    valueCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    currentRow += 1;
  });

  currentRow += 2;
  const noteCell = worksheet.getCell(currentRow, 1);
  noteCell.value =
    "Note: 'Program Only' indicates whether this export contains only program data (Yes) or includes process/execution data (No)";
  noteCell.font = { italic: true, size: 10 };
  noteCell.alignment = { wrapText: true, vertical: "top" };
  worksheet.mergeCells(currentRow, 1, currentRow, 2);
};

const createWorkbookWithDefaults = () => {
  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties = {
    fullCalcOnLoad: false,
  };
  return workbook;
};

const finalizeWorkbook = async (workbook) => {
  workbook.calcProperties = {
    fullCalcOnLoad: true,
  };
  return workbook.xlsx.writeBuffer();
};

const buildExerciseMap = (bundle) =>
  new Map((bundle.exercises || []).map((exercise) => [exercise._id, exercise]));

const buildProgramOnlyProcess = (bundle) => {
  const program = bundle.program;
  const process = {
    _id: "program-only-export",
    name: program.name,
    program_id: program._id,
    program_name: program.name,
    user_id: "",
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
    start_date: new Date().toISOString().split("T")[0],
    config: {
      weight_unit: "lbs",
      weight_rounding: 2.5,
      week_start_day: 0,
    },
    one_rm_profile: {
      enable_blockly_one_rm: false,
      squat: -1,
      bench: -1,
      deadlift: -1,
      blockly_one_rm: [],
    },
    exercise_recordings: [],
  };

  (bundle.program.custom_anchored_lifts || []).forEach((liftName) => {
    process.one_rm_profile[liftName.toLowerCase()] = -1;
  });

  (bundle.exercises || []).forEach((exercise) => {
    const block = program.blocks.find((candidate) => candidate._id === exercise.block_id);
    if (!block) {
      return;
    }
    process.exercise_recordings.push({
      exercise_id: exercise._id,
      weekly: Array.from({ length: block.duration }, () => ({
        sets: [],
        note: "",
      })),
    });
  });

  return process;
};

const addBlockWorksheets = (workbook, program, process, exerciseMap) => {
  program.blocks.forEach((block, blockIndex) => {
    const worksheetName = truncateWorksheetName(
      `Block ${blockIndex + 1} - ${block.name}`
    );
    const worksheet = workbook.addWorksheet(worksheetName);
    populateBlockWorksheet(
      worksheet,
      block,
      blockIndex,
      program,
      process,
      exerciseMap
    );
  });
};

export const exportProcessBundleToXlsxBuffer = async (bundle) => {
  const workbook = createWorkbookWithDefaults();
  const exerciseMap = buildExerciseMap(bundle);

  addBlockWorksheets(workbook, bundle.program, bundle.process, exerciseMap);

  const oneRmProfileWorksheet = workbook.addWorksheet("1RM Profile");
  create1RMProfileWorksheet(oneRmProfileWorksheet, bundle.process, bundle.program);

  const rpeChartWorksheet = workbook.addWorksheet("RPE Chart");
  createRPEChartWorksheet(rpeChartWorksheet);

  const metadataWorksheet = workbook.addWorksheet("Metadata");
  createMetadataWorksheet(metadataWorksheet, bundle.process, bundle.program);

  return {
    buffer: await finalizeWorkbook(workbook),
    filename: generateExportFilename("process", "xlsx", {
      name: bundle.process?.name || bundle.program?.name || "process",
    }),
  };
};

export const exportProgramBundleToXlsxBuffer = async (bundle) => {
  const workbook = createWorkbookWithDefaults();
  const mockProcess = buildProgramOnlyProcess(bundle);
  const exerciseMap = buildExerciseMap(bundle);

  addBlockWorksheets(workbook, bundle.program, mockProcess, exerciseMap);

  const oneRmProfileWorksheet = workbook.addWorksheet("1RM Profile");
  create1RMProfileWorksheet(oneRmProfileWorksheet, mockProcess, bundle.program);

  const rpeChartWorksheet = workbook.addWorksheet("RPE Chart");
  createRPEChartWorksheet(rpeChartWorksheet);

  const metadataWorksheet = workbook.addWorksheet("Metadata");
  createProgramMetadataWorksheet(metadataWorksheet, bundle.program);

  return {
    buffer: await finalizeWorkbook(workbook),
    filename: generateExportFilename("program", "xlsx", {
      name: bundle.program?.name || "program",
    }),
  };
};

export const exportBundleToXlsxBuffer = async (bundle) => {
  if (bundle?.process && bundle?.program && Array.isArray(bundle?.exercises)) {
    return exportProcessBundleToXlsxBuffer(bundle);
  }

  if (bundle?.program && Array.isArray(bundle?.exercises) && !bundle?.process) {
    return exportProgramBundleToXlsxBuffer(bundle);
  }

  throw new Error(
    "Unsupported bundle shape for XLSX export. Expected program-import or process-import."
  );
};
