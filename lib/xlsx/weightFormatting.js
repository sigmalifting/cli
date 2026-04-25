export const DEFAULT_WEIGHT_ROUNDING = 2.5;

export const getEffectiveWeightRounding = (config = {}) => {
  const rounding = Number(config?.weight_rounding);
  if (!Number.isFinite(rounding) || rounding <= 0) {
    return DEFAULT_WEIGHT_ROUNDING;
  }
  return rounding;
};

export const cleanRoundedValue = (value, roundingIncrement) => {
  if (value == null) {
    return value;
  }

  const incrementStr = String(roundingIncrement);
  const decimalIndex = incrementStr.indexOf(".");
  const numDecimalPlaces =
    decimalIndex === -1 ? 0 : incrementStr.length - decimalIndex - 1;

  return Number.parseFloat(value.toFixed(numDecimalPlaces));
};

export const formatWeight = (weight) => {
  if (weight === undefined || weight === null || weight < 0) {
    return "";
  }

  const normalized = Number.parseFloat(Number(weight).toFixed(6));
  return Number.isFinite(normalized) ? String(normalized) : "";
};
