// Export constants shared between Excel export and import functionality
export const TABLE_CONSTANTS = {
  EXERCISE_NAME_COL_WIDTH: 25,
  SET_NO_COL_WIDTH: 8,
  WEIGHT_COL_WIDTH: 12,
  REPS_COL_WIDTH: 8,
  RPE_COL_WIDTH: 8,
  ATTRIBUTES_COL_WIDTH: 20,
  NOTES_COL_WIDTH: 15,
  
  // Column indices within a day (0-based from day start)
  DAY_COLS: {
    SET_NO: 0,
    PRESC_WEIGHT: 1,
    PRESC_REPS: 2,
    PRESC_RPE: 3,
    PRESC_ATTRIBUTES: 4,
    PRESC_NOTES: 5,
    EXEC_WEIGHT: 6,
    EXEC_REPS: 7,
    EXEC_RPE: 8,
    EXEC_NOTES: 9
  },
  
  COLS_PER_DAY: 10,
  SPACING_BETWEEN_DAYS: 1,
  SPACING_BETWEEN_WEEKS: 2,
  HEADER_ROWS: 3
};