import { convertWeight } from "./weightConversions.js";
import { calculate1RM, calculateTrainingWeight } from "../rpe.js";

const DEBUG = false;
const log = DEBUG ? console.log : () => {};
const warn = DEBUG ? console.warn : () => {};

const getOneRm = (oneRmProfile, oneRmAnchor, blockIndex) => {
  if (!oneRmAnchor?.enabled || !oneRmAnchor.lift_type) {
    return undefined;
  }

  let baseRm;
  const liftType = oneRmAnchor.lift_type.toLowerCase();

  if (!oneRmProfile.enable_blockly_one_rm) {
    const value = oneRmProfile[liftType];
    if (typeof value === "number") {
      baseRm = value;
    } else {
      return undefined;
    }
  } else if (blockIndex !== undefined) {
    const blockProfile = oneRmProfile.blockly_one_rm?.[blockIndex];
    if (blockProfile) {
      const value = blockProfile[liftType];
      if (typeof value === "number") {
        baseRm = value;
      }
    }

    if (!baseRm || baseRm <= 0) {
      log(`Block-specific 1RM enabled but empty for ${liftType}`);
      return undefined;
    }
  }

  return baseRm && baseRm > 0 ? baseRm * (oneRmAnchor.ratio ?? 1.0) : undefined;
};

export const cleanRoundedValue = (value, roundingIncrement) => {
  if (value == null) {
    return value;
  }

  const incrementStr = String(roundingIncrement);
  const decimalIndex = incrementStr.indexOf(".");
  const numDecimalPlaces =
    decimalIndex === -1 ? 0 : incrementStr.length - decimalIndex - 1;

  return parseFloat(value.toFixed(numDecimalPlaces));
};

const applyWeightRounding = (weight, weightRounding) => {
  if (weightRounding <= 0) {
    return weight;
  }

  const rounded = Math.round(weight / weightRounding) * weightRounding;
  return cleanRoundedValue(rounded, weightRounding);
};

const calculateRpeBasedWeight = (
  referenceWeight,
  referenceReps,
  referenceRpe,
  targetRpe,
  targetReps,
  weightRounding
) => {
  if (
    !referenceReps ||
    !referenceRpe ||
    referenceReps <= 0 ||
    referenceRpe <= 0
  ) {
    warn("Missing reference data for RPE calculation");
    return undefined;
  }

  const estimated1RM = calculate1RM(
    referenceWeight,
    referenceReps,
    referenceRpe
  );
  if (estimated1RM <= 0) {
    warn("Failed to estimate 1RM");
    return undefined;
  }

  const adjustedRpe = Math.max(1, Math.min(10, targetRpe));

  const rawWeight = calculateTrainingWeight(
    estimated1RM,
    targetReps,
    adjustedRpe
  );
  if (rawWeight <= 0) {
    warn("Failed to calculate training weight");
    return -1;
  }

  return applyWeightRounding(rawWeight, weightRounding);
};

export const calculatePrescriptionForSet = (
  exerciseDefinition,
  weekIndex,
  setIndex,
  oneRmProfile,
  blockIndex,
  weightRounding,
  allSetsInExerciseInstance = [],
  processWeightUnit = "kg",
  totalWeeksInBlock,
  exerciseRecording
) => {
  const setGroups = exerciseDefinition.set_groups;
  if (!setGroups?.length) {
    warn("No set groups defined for this exercise.");
    return {};
  }

  let currentSetGroup = null;
  let setIndexWithinGroup = setIndex;
  let cumulativeSets = 0;

  for (const group of setGroups) {
    const numSetsThisWeek = group.weekly_num_sets?.[weekIndex];
    if (!numSetsThisWeek || numSetsThisWeek <= 0) {
      continue;
    }

    if (setIndex < cumulativeSets + numSetsThisWeek) {
      currentSetGroup = group;
      setIndexWithinGroup = setIndex - cumulativeSets;
      break;
    }
    cumulativeSets += numSetsThisWeek;
  }

  if (!currentSetGroup) {
    warn(`Set index ${setIndex} is out of bounds for week ${weekIndex}`);
    return {};
  }

  const variableParam = currentSetGroup.variable_parameter;
  let prescribedReps;
  let prescribedRpe;
  let calculatedWeight;

  if (variableParam !== "reps") {
    const repsValue = currentSetGroup.weekly_reps?.[weekIndex];
    if (repsValue !== undefined && repsValue >= 0) {
      prescribedReps = repsValue;
    }
  }

  if (variableParam !== "rpe") {
    const rpeValue = currentSetGroup.weekly_rpe?.[weekIndex];
    if (rpeValue !== undefined && rpeValue >= 0) {
      prescribedRpe = rpeValue;
    }
  }

  const deloadConfig = exerciseDefinition.deload_config;
  const isDeloadWeek =
    deloadConfig?.enabled &&
    totalWeeksInBlock !== undefined &&
    weekIndex === totalWeeksInBlock - 1;

  if (isDeloadWeek && exerciseRecording && weekIndex > 0) {
    const previousWeekIndex = weekIndex - 1;
    const previousWeekData = exerciseRecording.weekly?.[previousWeekIndex];

    if (previousWeekData?.sets?.length > 0) {
      let prevWeekCumulativeSets = 0;

      for (const group of setGroups) {
        if (group.group_id === currentSetGroup.group_id) {
          const prevWeekNumSetsInGroup =
            group.weekly_num_sets?.[previousWeekIndex] || 0;
          const targetSetIndexWithinGroup = Math.min(
            setIndexWithinGroup,
            prevWeekNumSetsInGroup - 1
          );

          if (targetSetIndexWithinGroup >= 0 && prevWeekNumSetsInGroup > 0) {
            const absolutePrevWeekSetIndex =
              prevWeekCumulativeSets + targetSetIndexWithinGroup;

            if (absolutePrevWeekSetIndex < previousWeekData.sets.length) {
              const previousSet =
                previousWeekData.sets[absolutePrevWeekSetIndex];

              if (
                previousSet.weight !== undefined &&
                previousSet.weight > 0 &&
                previousSet.completed
              ) {
                const deloadPercentage = deloadConfig.percentage || 85;
                const rawDeloadWeight =
                  previousSet.weight * (deloadPercentage / 100);
                calculatedWeight = applyWeightRounding(
                  rawDeloadWeight,
                  weightRounding
                );
              }
            }
          }
          break;
        }
        prevWeekCumulativeSets +=
          group.weekly_num_sets?.[previousWeekIndex] || 0;
      }
    }

    if (variableParam === "weight") {
      prescribedRpe = undefined;
    }
  } else if (variableParam === "weight") {
    calculatedWeight = undefined;
  } else if (currentSetGroup.backoff_config?.enabled) {
    const backoffConfig = currentSetGroup.backoff_config;
    const dependsOnGroupId = backoffConfig.depends_on_set_group_id;

    if (dependsOnGroupId) {
      let referenceWeight;
      let referenceSet;
      let cumulative = 0;
      let endIndex = -1;

      for (const grp of setGroups) {
        const sets = grp.weekly_num_sets?.[weekIndex] ?? 0;
        if (grp.group_id === dependsOnGroupId) {
          endIndex = cumulative + sets - 1;
          break;
        }
        cumulative += sets;
      }

      if (endIndex >= 0 && endIndex < allSetsInExerciseInstance.length) {
        referenceSet = allSetsInExerciseInstance[endIndex];
        referenceWeight = referenceSet.weight;
      }

      if (referenceWeight !== undefined && referenceWeight >= 0) {
        if (backoffConfig.type === "%_of_weight") {
          const backoffValueThisWeek =
            backoffConfig.weekly_percentage?.[weekIndex];
          if (backoffValueThisWeek !== undefined) {
            const rawBackoffWeight =
              referenceWeight * (backoffValueThisWeek / 100);
            calculatedWeight = applyWeightRounding(
              rawBackoffWeight,
              weightRounding
            );
          }
        } else if (backoffConfig.type === "target_rpe") {
          const targetRpe = backoffConfig.weekly_rpe?.[weekIndex];
          if (targetRpe && referenceSet) {
            const backoffReps =
              currentSetGroup.weekly_reps?.[weekIndex] ||
              referenceSet.reps ||
              0;
            calculatedWeight = calculateRpeBasedWeight(
              referenceWeight,
              referenceSet.reps,
              referenceSet.rpe,
              targetRpe,
              backoffReps,
              weightRounding
            );
          }
        }
      }
    }
  } else {
    const oneRm = getOneRm(
      oneRmProfile,
      exerciseDefinition.one_rm_anchor,
      blockIndex
    );
    log(`Standard calc - oneRm: ${oneRm}, blockIndex: ${blockIndex}`);

    if (currentSetGroup.mix_weight_config?.enabled) {
      const mixConfig = currentSetGroup.mix_weight_config;
      const weekPercentage = mixConfig.weekly_weight_percentage?.[weekIndex];
      let absoluteWeight = mixConfig.weekly_weight_absolute?.[weekIndex];

      if (
        oneRm !== undefined &&
        weekPercentage !== undefined &&
        absoluteWeight !== undefined
      ) {
        if (
          mixConfig.weight_unit &&
          mixConfig.weight_unit !== processWeightUnit
        ) {
          absoluteWeight = convertWeight(
            absoluteWeight,
            mixConfig.weight_unit,
            processWeightUnit
          );
        }

        const percentageWeight = oneRm * (weekPercentage / 100);
        const finalWeight = percentageWeight + absoluteWeight;
        calculatedWeight = applyWeightRounding(finalWeight, weightRounding);
      }
    } else if (
      currentSetGroup.weekly_weight_percentage?.[weekIndex] !== undefined &&
      oneRm !== undefined
    ) {
      const percentage = currentSetGroup.weekly_weight_percentage[weekIndex];
      const rawWeight = oneRm * (percentage / 100);
      calculatedWeight = applyWeightRounding(rawWeight, weightRounding);
    }
  }

  const fatigueDropConfig = currentSetGroup.fatigue_drop_config;

  if (!isDeloadWeek && fatigueDropConfig?.enabled) {
    const rpeCap = fatigueDropConfig.rpe_cap?.[weekIndex];

    if (rpeCap !== undefined && rpeCap !== null) {
      let startIndex = 0;
      let cumulative = 0;

      for (const grp of setGroups) {
        if (grp.group_id === currentSetGroup.group_id) {
          startIndex = cumulative;
          break;
        }
        cumulative += grp.weekly_num_sets?.[weekIndex] ?? 0;
      }

      for (
        let index = startIndex;
        index < setIndex && index < allSetsInExerciseInstance.length;
        index += 1
      ) {
        const prevSet = allSetsInExerciseInstance[index];

        if (
          prevSet.completed &&
          prevSet.rpe !== undefined &&
          prevSet.rpe >= rpeCap
        ) {
          const referenceWeight = prevSet.weight;

          if (referenceWeight !== undefined && referenceWeight > 0) {
            if (fatigueDropConfig.type === "%_of_weight") {
              const dropPercentage =
                fatigueDropConfig.weekly_percentage?.[weekIndex];
              if (dropPercentage !== undefined) {
                const rawDroppedWeight =
                  referenceWeight * (dropPercentage / 100);
                calculatedWeight = applyWeightRounding(
                  rawDroppedWeight,
                  weightRounding
                );
                log(
                  `Applied fatigue drop: ${dropPercentage}% of ${referenceWeight} = ${calculatedWeight}`
                );
              }
            } else if (fatigueDropConfig.type === "target_rpe") {
              const targetRpe = fatigueDropConfig.weekly_rpe?.[weekIndex];
              if (targetRpe) {
                const currentReps =
                  currentSetGroup.weekly_reps?.[weekIndex] || prevSet.reps || 0;
                calculatedWeight = calculateRpeBasedWeight(
                  referenceWeight,
                  prevSet.reps,
                  prevSet.rpe,
                  targetRpe,
                  currentReps,
                  weightRounding
                );
              }
            }
          }
          break;
        }
      }
    }
  }

  log("Result:", {
    weight: calculatedWeight,
    reps: prescribedReps,
    rpe: prescribedRpe,
  });

  return {
    weight: calculatedWeight,
    reps: prescribedReps,
    rpe: prescribedRpe,
  };
};
