export const KG_TO_LBS = 2.20462;
export const LBS_TO_KG = 0.453592;

export const kgToLbs = (kg) => kg * KG_TO_LBS;

export const lbsToKg = (lbs) => lbs * LBS_TO_KG;

export const convertWeight = (weight, fromUnit, toUnit) => {
  if (fromUnit === toUnit) {
    return weight;
  }

  if (fromUnit === "kg" && toUnit === "lbs") {
    return kgToLbs(weight);
  }

  return lbsToKg(weight);
};

