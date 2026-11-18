import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  Container,
  Grid,
  Paper,
  Typography,
  Button,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Box,
  Alert,
  Divider,
  Card,
  CardContent,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Chip,
  LinearProgress,
  Radio, // <-- Add
  RadioGroup, // <-- Add
  FormControlLabel, // <-- Add
  FormLabel,
} from "@mui/material";
import { read, utils } from "xlsx";
import {
  BarChart,
  Bar,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  LabelList,
} from "recharts";
import {
  MapContainer,
  TileLayer,
  CircleMarker,
  Circle,
  Popup,
  Tooltip as LeafletTooltip,
  useMap,
  LayersControl,
  LayerGroup,
} from "react-leaflet";
import "leaflet/dist/leaflet.css";
import * as math from "mathjs";
import L from "leaflet";

const NBHDMODIFIERTHRESHOLD = 0.05;
// IAAO Standard: Calculates Median ASR and Coefficient of Dispersion (COD)
function computeRatioStats(ratioArray) {
  const valid = ratioArray.filter(
    (r) => typeof r === "number" && !isNaN(r) && r > 0,
  );
  if (valid.length === 0) {
    return { median: 1.0, cod: 0.0, count: 0 };
  }

  const sorted = [...valid].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const median =
    sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;

  if (median === 0) return { median: 0, cod: 0, count: sorted.length };

  // COD = (Mean Absolute Deviation from Median / Median) * 100%
  const meanAbsDev =
    sorted.reduce((sum, val) => sum + Math.abs(val - median), 0) /
    sorted.length;
  const cod = (meanAbsDev / median) * 100;

  return {
    median: parseFloat(median.toFixed(4)),
    cod: parseFloat(cod.toFixed(2)),
    count: sorted.length,
  };
}

function formatExcelDate(val) {
  if (!val) return null;
  if (typeof val === "number") {
    // Convert Excel serial date to YYYY-MM-DD
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    return date.toISOString().split("T")[0];
  }
  return String(val);
}
// ==========================================
// 1. ADVANCED ECONOMETRIC MODEL CLASS
// ==========================================
export class EdmontonGravityExposureModel {
  constructor(eval_year = 2026, ring_size = 0, decay_exponent = 1.5) {
    this.typologies = ["One-Unit", "Two-Unit", "Multi-Unit", "FivePlus-Unit"];
    this.decay_exponent = decay_exponent;
    this.ring_size = ring_size;
    this.eval_year = eval_year;
    this.model = null; // Stores coefficients, formula, intercepts
    this.exposure_cols = this.typologies.map((t) => `exposure_${t}`);
    this.allSales = [];
    this.infillData = [];
    this.propertyInventory = [];
  }

  // --- 1. Load Study Data from InfillStudyData.xlsx ---
  load_study_data(rawSales, rawInfills, evalYear = 2026) {
    this.eval_year = evalYear;
    this.allSales = rawSales.map((row) => {
      const actYear = row.ACT_YEARBUILT || 1959;
      const effYear = row.EFF_YEARBUILT || 1961;
      return {
        id: row.ID,
        rollNbr: row.ROLL_NBR,
        nbhdName: row.NBHD_NAME,
        price: row.PRICE,
        adjPrice: row.ADJUSTED_PRICE,
        totalAsmt: row.TOTAL_AP_ASMNT,
        unadjUnitPriceSm: row.UNADJUSTED_UNIT_PRICE,
        unitPriceSm: row.UNIT_PRICE || 0,
        saleDate: formatExcelDate(row.SALE_DATE),
        soldYear: row.SOLD_YEAR || 0,
        lotSizeSm: row.LOTSIZE || 0,
        flrAreaSm: row.FLR_AREA || 0,
        actYearBuilt: actYear,
        effYearBuilt: effYear,
        age: evalYear - actYear + 1,
        quality: row.QUALITY || 4,
        condition: row.CONDITION || 4,
        propX: row.PROP_X, // Latitude
        propY: row.PROP_Y, // Longitude
        totalUnitsAdded: 0,
        ringSize: 0,
        includeFlag: row.SOLD_YEAR - actYear > 2 ? "Y" : "N",
        predUnitPrice: 0,
      };
    });

    this.infillData = rawInfills.map((row) => ({
      infillId: row.INFILL_ID,
      nbhdName: row.NBHD_NAME,
      typology: row.TYPOLOGY,
      permitDate: formatExcelDate(row.PERMIT_DATE),
      permitYear: row.PERMIT_YEAR,
      completionYear: row.COMPLETION_YEAR,
      unitsAdded: row.UNITS_ADDED || 1,
      infillX: row.INFILL_X, // Latitude
      infillY: row.INFILL_Y, // Longitude
      includeFlag: "Y",
    }));

    return { allSales: this.allSales, infillData: this.infillData };
  }

  // --- 2. Load External Property Inventory (Feature 2) ---
  load_property_inventory(rawInventory, evalYear = null) {
    const targetYear = evalYear || this.eval_year;
    this.propertyInventory = rawInventory.map((row) => {
      const effYear = row.EFF_YEARBUILT || 1970;
      return {
        rollNbr: row.ROLL_NBR,
        totalAsmt: row.TOTAL_AP_ASMNT,
        nbhdName: row.NBHD_NAME || "UNKNOWN",
        lotSizeSm: row.LOTSIZE || 500,
        flrAreaSm: row.FLR_AREA || 150,
        effYearBuilt: effYear,
        age: targetYear - effYear + 1,
        quality: row.QUALITY || 4,
        condition: row.CONDITION || 3,
        propX: row.LATITUDE,
        propY: row.LONGITUDE,
        predUnitPrice: 0,
        predTotalValue: 0,
        exposures: {},
      };
    });
    return this.propertyInventory;
  }

  // --- 3. Load Previously Saved Model JSON (Feature 1) ---
  load_model(modelData) {
    if (!modelData || !modelData.coefficients) {
      throw new Error("Invalid model format. Missing coefficients.");
    }
    this.model = modelData;
    this.eval_year = modelData.eval_year || this.eval_year;
    this.ring_size = modelData.ring_size || this.ring_size;
    this.decay_exponent = modelData.decay_exponent || this.decay_exponent;
    this.typologies = modelData.typologies || this.typologies;
    this.exposure_cols = modelData.exposure_cols || this.exposure_cols;
    this.nbhdFixedEffects = modelData.nbhdFixedEffects || {};
    this.smearingFactor = modelData.smearingFactor || 1.0;
    this.nbhdMedianASR = modelData.nbhdMedianASR || {};
    this.nbhdModifiers = modelData.nbhdModifiers || {};
    this.calibrationStats = modelData.calibrationStats || null;
    return this.model;
  }

  // Haversine distance in meters
  _calculateDistance(lat1, lon1, lat2, lon2) {
    const R = 6371e3;
    const p1 = (lat1 * Math.PI) / 180;
    const p2 = (lat2 * Math.PI) / 180;
    const dp = ((lat2 - lat1) * Math.PI) / 180;
    const dl = ((lon2 - lon1) * Math.PI) / 180;
    const a =
      Math.sin(dp / 2) * Math.sin(dp / 2) +
      Math.cos(p1) * Math.cos(p2) * Math.sin(dl / 2) * Math.sin(dl / 2);
    return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
  }

  // Continuous distance-decay gravity exposure calculation
  compute_gravity_exposure(properties, infills) {
    const activeProperties = properties.filter((p) => p.includeFlag === "Y");
    const activeInfills = infills.filter((i) => i.includeFlag === "Y");

    activeProperties.forEach((prop) => {
      this.typologies.forEach((t) => {
        prop[`exposure_${t}`] = 0.0;
      });

      activeInfills.forEach((infill) => {
        const dist = this._calculateDistance(
          prop.propX,
          prop.propY,
          infill.infillX,
          infill.infillY,
        );
        if (dist <= this.ring_size) {
          const weight =
            infill.unitsAdded / Math.pow(1 + dist, this.decay_exponent);
          const colKey = `exposure_${infill.typology}`;
          if (prop[colKey] !== undefined) {
            prop[colKey] += weight;
          }
        }
      });
    });
    return activeProperties;
  }
  //getMean
  _getMean = (numArr) => {
    if (numArr.length === 0) return 0; // Prevents dividing by zero
    const sum = numArr.reduce((acc, val) => acc + val, 0);
    return sum / numArr.length;
  };

  //getMedian
  _getMedian = (numArr) => {
    const sorted = [...numArr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    const median =
      sorted.length % 2 !== 0
        ? sorted[mid]
        : (sorted[mid - 1] + sorted[mid]) / 2;

    return median;
  };

  // Multi-way Fixed Effects Absorption (Method of Alternating Projections)
  _absorb_fixed_effects(data, targetCols, groupCols, maxIter = 40, tol = 1e-5) {
    let demeaned = data.map((row) => ({ ...row }));

    for (let iter = 0; iter < maxIter; iter++) {
      let maxShift = 0;
      for (const groupCol of groupCols) {
        const stats = {};
        demeaned.forEach((row) => {
          const g = row[groupCol];
          if (!stats[g]) stats[g] = { count: 0, sums: {} };
          stats[g].count += 1;
          targetCols.forEach((col) => {
            stats[g].sums[col] = (stats[g].sums[col] || 0) + row[col];
          });
        });

        demeaned.forEach((row) => {
          const g = row[groupCol];
          targetCols.forEach((col) => {
            const mean = stats[g].sums[col] / stats[g].count;
            row[col] -= mean;
            if (Math.abs(mean) > maxShift) maxShift = Math.abs(mean);
          });
        });
      }
      if (maxShift < tol) break;
    }
    return demeaned;
  }

  // --- 4. Model Fit via OLS Matrix Inversion ---
  fit(properties, infills, benchmarkProperty = null) {
    const data = this.compute_gravity_exposure(properties, infills);

    if (!data || data.length === 0) {
      throw new Error(`No sales found for evaluation year ${this.eval_year}.`);
    }

    data.forEach((row) => {
      row.log_unit_price = Math.log(row.unitPriceSm > 0 ? row.unitPriceSm : 1);
      row.log_living_area = Math.log(row.flrAreaSm > 0 ? row.flrAreaSm : 1);
      row.log_lot_size = Math.log(row.lotSizeSm > 0 ? row.lotSizeSm : 1);
      row.log_age = Math.log(row.age > 0 ? row.age : 1);
    });

    const controls = [
      "log_living_area",
      "log_lot_size",
      "log_age",
      "quality",
      "condition",
    ];
    const X_cols = [...this.exposure_cols, ...controls];

    // Compute baseline means for intercept recovery and sample means for typical property
    const meanY = math.mean(data.map((d) => d.log_unit_price));
    const meanX = {};
    const sampleAverages = {
      salesCount: data.length,
      typicalLivingArea: Math.round(
        Math.exp(this._getMedian(data.map((d) => d.log_living_area))),
      ),
      typicalLotSize: Math.round(
        Math.exp(this._getMedian(data.map((d) => d.log_lot_size))),
      ),
      typicalAge: Math.round(
        Math.exp(this._getMedian(data.map((d) => d.log_age))),
      ),
      typicalQuality: this._getMedian(data.map((d) => d.quality)).toFixed(2),
      typicalCondition: this._getMedian(data.map((d) => d.condition)).toFixed(
        2,
      ),
      typicalUnitPrice: Math.round(Math.exp(meanY)),
    };
    // X_cols.forEach((col) => {
    //   meanX[col] = math.mean(data.map((d) => d[col]));
    // });

    // Absorb Neighbourhood Fixed Effects
    const demeanedData = this._absorb_fixed_effects(
      data,
      [...X_cols, "log_unit_price"],
      ["nbhdName"],
    );

    try {
      const Y = demeanedData.map((r) => [r.log_unit_price]);
      const X = demeanedData.map((r) => X_cols.map((c) => r[c]));

      const X_t = math.transpose(X);
      let X_t_X = math.multiply(X_t, X);
      // Convert to native array if returned as mathjs Matrix
      if (X_t_X._data) X_t_X = X_t_X._data;
      else if (X_t_X.toArray) X_t_X = X_t_X.toArray();

      // RIDGE REGULARIZATION (lambda * I):
      // Prevents matrix singularity when an infill typology has 0 exposures in this ring/year
      const p = X_cols.length;
      const lambda = 1e-5;
      for (let i = 0; i < p; i++) {
        X_t_X[i][i] = (X_t_X[i][i] || 0) + lambda;
      }

      const X_t_X_inv = math.inv(X_t_X);
      const Beta = math.multiply(X_t_X_inv, math.multiply(X_t, Y));

      const coefficients = {};
      X_cols.forEach((col, idx) => {
        coefficients[col] = Beta[idx][0];
      });
      console.log(coefficients);
      // ========================================================
      // CALIBRATION 1: RECOVER INDIVIDUAL NEIGHBOURHOOD INTERCEPTS
      // ========================================================
      const nbhdGroups = {};
      data.forEach((row) => {
        const g = row.nbhdName;
        if (!nbhdGroups[g]) {
          nbhdGroups[g] = { count: 0, sumY: 0, sumX: {} };
          X_cols.forEach((c) => {
            nbhdGroups[g].sumX[c] = 0;
          });
        }
        nbhdGroups[g].count += 1;
        nbhdGroups[g].sumY += row.log_unit_price;
        X_cols.forEach((c) => {
          nbhdGroups[g].sumX[c] += row[c];
        });
      });

      const nbhdFixedEffects = {};
      Object.keys(nbhdGroups).forEach((g) => {
        const grp = nbhdGroups[g];
        const meanY_g = grp.sumY / grp.count;
        let alpha_g = meanY_g;
        X_cols.forEach((col) => {
          const meanX_col = grp.sumX[col] / grp.count;
          alpha_g -= coefficients[col] * meanX_col;
        });
        nbhdFixedEffects[g] = alpha_g;
      });

      // Fallback Grand Mean Intercept (for unobserved neighbourhoods)
      const grandMeanY = math.mean(data.map((d) => d.log_unit_price));
      let baseIntercept = grandMeanY;
      X_cols.forEach((col) => {
        baseIntercept -= coefficients[col] * math.mean(data.map((d) => d[col]));
      });

      // ========================================================
      // CALIBRATION 2: DUAN'S SMEARING FACTOR & RATIO METRICS
      // ========================================================
      let sumExpResiduals = 0;

      data.forEach((row) => {
        const intercept = nbhdFixedEffects[row.nbhdName] || baseIntercept;
        let predLog = intercept;
        X_cols.forEach((col) => {
          predLog += coefficients[col] * row[col];
        });

        const residual = row.log_unit_price - predLog;
        sumExpResiduals += Math.exp(residual);
      });

      // Smearing factor corrects log transformation bias
      const smearingFactor = sumExpResiduals / data.length;

      // ========================================================
      // 3. IN-SAMPLE PREDICTION & RAW  (ASR) per sale
      // ========================================================
      const nbhdAsrBuckets = {};

      data.forEach((row) => {
        // console.log(row);
        const intercept = nbhdFixedEffects[row.nbhdName] || baseIntercept;
        let predLog = intercept;
        X_cols.forEach((col) => {
          predLog += coefficients[col] * row[col];
        });

        // Predicted unit price and total value for this sale
        row.basePredUnitPrice = Math.exp(predLog) * smearingFactor;
        row.basePredTotalValue = row.basePredUnitPrice * row.flrAreaSm;

        // ASR = TOTAL_AP_ASMNT / Predicted Total Value
        // if (row.totalAsmt > 0 && row.basePredTotalValue > 0) {
        if (row.totalAsmt > 0 && row.adjPrice > 0) {
          // row.rawPredASR = row.totalAsmt / row.basePredTotalValue;
          row.rawASR = row.totalAsmt / row.adjPrice;
          if (!nbhdAsrBuckets[row.nbhdName]) {
            nbhdAsrBuckets[row.nbhdName] = [];
          }
          nbhdAsrBuckets[row.nbhdName].push(row.rawASR);
        } else {
          row.rawASR = null;
        }
      });

      // console.log(nbhdAsrBuckets);
      // ========================================================
      // 4. NEIGHBOURHOOD MODIFIER CALCULATION
      // Rule: If |1.00 - Median(ASR)| >= 0.3, Modifier = 1 / Median(ASR)
      //       Else, Modifier = 1.00
      // ========================================================
      const nbhdMedianASR = {};
      const nbhdMeanASR = {};
      const nbhdModifiers = {};
      const modifiedNbhdsCount = {
        adjusted: 0,
        total: Object.keys(nbhdGroups).length,
      };

      Object.keys(nbhdGroups).forEach((g) => {
        const asrs = nbhdAsrBuckets[g] || [];
        if (asrs.length > 0) {
          const sorted = [...asrs].sort((a, b) => a - b);
          const mid = Math.floor(sorted.length / 2);
          const median =
            sorted.length % 2 !== 0
              ? sorted[mid]
              : (sorted[mid - 1] + sorted[mid]) / 2;

          const mean = this._getMean(sorted);

          nbhdMedianASR[g] = parseFloat(median.toFixed(4));
          nbhdMeanASR[g] = parseFloat(mean.toFixed(4));

          const diff = Math.abs(1.0 - median); //switch this to mean if required.
          if (diff >= NBHDMODIFIERTHRESHOLD) {
            // Setting Modifier = meanASR scales predicted value so that:
            // New ASR = Assessment / (Pred * Modifier) = ASR / meanASR = 1.00
            nbhdModifiers[g] = parseFloat(median).toFixed(4); //switch this to mean if required.
            modifiedNbhdsCount.adjusted += 1;
          } else {
            nbhdModifiers[g] = 1.0;
          }
        } else {
          nbhdMedianASR[g] = 1.0;
          nbhdMeanASR[g] = 1.0;
          nbhdModifiers[g] = 1.0;
        }
        // console.log(
        //   `${g}: MeanASR: ${nbhdMeanASR[g].toFixed(3)} - MedianASR: ${nbhdMedianASR[g].toFixed(3)}`,
        // );
      });

      // ========================================================
      // 5. COMPUTE FINAL ASR, FINAL MEDIAN ASR, AND FINAL COD
      // ========================================================
      const rawAsrList = [];
      const finalAsrList = [];

      data.forEach((row) => {
        const intercept = nbhdFixedEffects[row.nbhdName] || baseIntercept;
        let predLog = intercept;
        X_cols.forEach((col) => {
          predLog += coefficients[col] * row[col];
        });
        const unadjustedPred = Math.exp(predLog);
        row.rawPredASR = unadjustedPred / row.unitPriceSm;

        if (row.rawPredASR !== null && row.rawPredASR > 0) {
          const mod = nbhdModifiers[row.nbhdName] || 1.0;
          // Applying the modifier scales the ratio: finalASR = rawASR * median -> median = 1.00
          row.finalASR = row.rawPredASR * mod;

          rawAsrList.push(row.rawASR);
          finalAsrList.push(row.finalASR);
        }
      });

      // Calculate overall calibration statistics (Before vs After)
      const rawRatioStats = computeRatioStats(rawAsrList);
      const finalRatioStats = computeRatioStats(finalAsrList);

      this.model = {
        eval_year: this.eval_year,
        ring_size: this.ring_size,
        decay_exponent: this.decay_exponent,
        typologies: this.typologies,
        exposure_cols: this.exposure_cols,
        intercept: baseIntercept,
        nbhdFixedEffects: nbhdFixedEffects,
        smearingFactor: smearingFactor,
        nbhdMedianASR: nbhdMedianASR, // Raw Median ASRs
        nbhdModifiers: nbhdModifiers, // 1/median modifiers
        modifierStats: modifiedNbhdsCount,
        calibrationStats: {
          rawMedianASR: rawRatioStats.median,
          rawCOD: rawRatioStats.cod,
          finalMedianASR: finalRatioStats.median,
          finalCOD: finalRatioStats.cod,
          sampleCount: rawRatioStats.count,
        },
        sampleAverages: sampleAverages,
        coefficients: coefficients,
        formula: `ln(UnitPrice) = α[nbhd] + ${Object.entries(coefficients)
          .map(([k, v]) => `(${v.toFixed(4)})*${k}`)
          .join(" + ")}`,
      };

      return this.model;
    } catch (err) {
      console.error("Matrix estimation error:", err);
      return null;
    }
  }

  // --- 5. Out-of-Sample Prediction on Property Inventory, Calibrated Intercept Applied ---
  predict(rollNbr, inventory, infillData) {
    if (!this.model) throw new Error("No model loaded or trained.");
    const prop = inventory.find((p) => String(p.rollNbr) === String(rollNbr));
    if (!prop) throw new Error(`Property ${rollNbr} not found in inventory.`);

    // 1. Calculate gravity exposures dynamically against active infills
    const activeInfills = infillData.filter((i) => i.includeFlag === "Y");
    const exposures = {};
    this.typologies.forEach((t) => {
      exposures[t] = 0.0;
    });

    activeInfills.forEach((infill) => {
      const dist = this._calculateDistance(
        prop.propX,
        prop.propY,
        infill.infillX,
        infill.infillY,
      );
      if (dist <= this.ring_size) {
        const weight =
          infill.unitsAdded / Math.pow(1 + dist, this.decay_exponent);
        if (exposures[infill.typology] !== undefined) {
          exposures[infill.typology] += weight;
        }
      }
    });

    prop.exposures = exposures;

    // 2. Compute log predictors
    const log_living_area = Math.log(prop.flrAreaSm > 0 ? prop.flrAreaSm : 1);
    const log_lot_size = Math.log(prop.lotSizeSm > 0 ? prop.lotSizeSm : 1);
    const log_age = Math.log(prop.age > 0 ? prop.age : 1);

    const featureVector = {
      log_living_area: log_living_area,
      log_lot_size: log_lot_size,
      log_age: log_age,
      quality: prop.quality,
      condition: prop.condition,
    };
    this.typologies.forEach((t) => {
      featureVector[`exposure_${t}`] = exposures[t];
    });

    // 3. CALIBRATION APPLICATION:
    // Look up neighbourhood fixed effect, fallback to city grand mean if new neighbourhood
    const nbhdIntercept =
      (this.model.nbhdFixedEffects &&
        this.model.nbhdFixedEffects[prop.nbhdName]) ||
      this.model.intercept ||
      0;

    let logPredPrice = nbhdIntercept;
    let logNoInfillPrice = nbhdIntercept; // Counterfactual: all exposures = 0

    Object.keys(this.model.coefficients).forEach((col) => {
      if (featureVector[col] !== undefined) {
        const contribution = this.model.coefficients[col] * featureVector[col];
        logPredPrice += contribution;

        // Only accumulate non-exposure structural variables for the zero-infill baseline
        if (!col.startsWith("exposure_")) {
          logNoInfillPrice += contribution;
        }
      }
    });

    // Apply Duan's smearing factor to remove geometric log bias
    const smearing = this.model.smearingFactor || 1.0;
    const baseUnitPrice = Math.exp(logPredPrice) * smearing;
    const baseNoInfillUnitPrice = Math.exp(logNoInfillPrice) * smearing;

    // 4. Retrieve Neighbourhood Modifier & Raw Median ASR
    const modifier =
      (this.model.nbhdModifiers && this.model.nbhdModifiers[prop.nbhdName]) ||
      1.0;
    const rawMedianASR =
      (this.model.nbhdMedianASR && this.model.nbhdMedianASR[prop.nbhdName]) ||
      1.0;

    // Calibrated valuation: scaling denominator by rawMedian (division by 1/median)
    prop.basePredUnitPrice = baseUnitPrice;
    prop.basePredTotalValue = baseUnitPrice * prop.flrAreaSm;
    prop.predUnitPrice =
      modifier !== 1.0 ? baseUnitPrice * modifier : baseUnitPrice;
    prop.predTotalValue = prop.predUnitPrice * prop.flrAreaSm;

    // Counterfactual (0 Gravity Exposure) properties (also persisted for CSV export)
    prop.noInfillUnitPrice =
      modifier !== 1.0
        ? baseNoInfillUnitPrice * modifier
        : baseNoInfillUnitPrice;
    prop.noInfillTotalValue = prop.noInfillUnitPrice * prop.flrAreaSm;
    prop.infillImpactDollar = prop.predTotalValue - prop.noInfillTotalValue;

    const infillImpactPct =
      prop.noInfillTotalValue > 0
        ? (prop.infillImpactDollar / prop.noInfillTotalValue) * 100
        : 0;

    return {
      rollNbr: prop.rollNbr,
      nbhdName: prop.nbhdName,
      nbhdIntercept: nbhdIntercept,
      flrAreaSm: prop.flrAreaSm,
      lotSizeSm: prop.lotSizeSm,
      age: prop.age,
      exposures: exposures,
      // rawMedianASR: rawMedianASR,
      // finalMedianASR: modifier !== 1.00 ? rawMedianASR * modifier : rawMedianASR,
      nbhdModifier: modifier,
      basePredTotalValue: prop.basePredTotalValue,
      predUnitPrice: prop.predUnitPrice,
      predTotalValue: prop.predTotalValue,
      noInfillUnitPrice: prop.noInfillUnitPrice,
      noInfillTotalValue: prop.noInfillTotalValue,
      infillImpactDollar: prop.infillImpactDollar,
      infillImpactPct: infillImpactPct,
      assessedValue: prop.totalAsmt,
    };
  }

  get_full_regression_table() {
    if (!this.model) return [];
    return Object.keys(this.model.coefficients).map((k) => {
      let pctImpact = 0;
      if (k.includes("log_"))
        pctImpact = (this.model.coefficients[k] / this.model.intercept) * 100;
      if (k === "quality" || k === "condition")
        pctImpact =
          ((this.model.coefficients[k] *
            Math.log(this.model.sampleAverages.typicalQuality)) /
            this.model.intercept) *
          100;
      if (k.includes("exposure_"))
        pctImpact = (this.model.coefficients[k] / this.model.intercept) * 100;
      return {
        variable: k,
        coefficient: this.model.coefficients[k],
        pct_impact: pctImpact,
      };
    });
  }
}

// Map Auto-Centering Hook
// function MapFocus({ center }) {
//   const map = useMap();
//   useEffect(() => {
//     if (center) map.flyTo(center, 14);
//   }, [center, map]);
//   return null;
// }

// Haversine distance helper (returns distance in meters)
export function calculateDistanceMeters(lat1, lon1, lat2, lon2) {
  const R = 6371e3; // Earth radius in meters
  const p1 = (lat1 * Math.PI) / 180;
  const p2 = (lat2 * Math.PI) / 180;
  const dp = ((lat2 - lat1) * Math.PI) / 180;
  const dl = ((lon2 - lon1) * Math.PI) / 180;
  const a =
    Math.sin(dp / 2) * Math.sin(dp / 2) +
    Math.cos(p1) * Math.cos(p2) * Math.sin(dl / 2) * Math.sin(dl / 2);
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function MapViewController({
  selectedNbhd,
  filteredSales,
  filteredInfills,
  subjectProp,
  ringSize,
  focusMode, // 'property' or 'neighbourhood'
}) {
  const map = useMap();

  // 1. Maintain canvas size integrity on window resize
  useEffect(() => {
    const handleResize = () => map.invalidateSize();
    window.addEventListener("resize", handleResize);
    const timer = setTimeout(() => map.invalidateSize(), 250);
    return () => {
      window.removeEventListener("resize", handleResize);
      clearTimeout(timer);
    };
  }, [map]);

  // 2. Camera controller: Priority to Subject Property Buffer, fallback to Neighbourhood
  useEffect(() => {
    if (focusMode === "property" && subjectProp) {
      const center = L.latLng(subjectProp.propX, subjectProp.propY);

      // Compute bounding box covering the full buffer ring diameter + padding margin
      const bufferDiameter = Math.max(ringSize || 50, 30) * 2.3;
      const ringBounds = center.toBounds(bufferDiameter);

      map.flyToBounds(ringBounds, {
        duration: 1.4,
        padding: [30, 30],
        maxZoom: 18, // Allows high-detail street level inspection
      });
      return;
    }

    // 3. Fallback: Zoom to Neighbourhood or City View
    if (selectedNbhd === "All") {
      map.flyTo([53.5461, -113.4938], 11, { duration: 1.2 });
      return;
    }

    const coords = [
      ...filteredSales.map((s) => [s.propX, s.propY]),
      ...filteredInfills.map((i) => [i.infillX, i.infillY]),
    ];

    if (coords.length > 0) {
      const bounds = L.latLngBounds(coords);
      map.flyToBounds(bounds, {
        padding: [40, 40],
        maxZoom: 15,
        duration: 1.2,
      });
    }
  }, [
    subjectProp,
    ringSize,
    focusMode,
    selectedNbhd,
    filteredSales,
    filteredInfills,
    map,
  ]);

  return null;
}

// Automatically recalculates Leaflet tile boundaries on screen resize
function MapResponsiveHandler({ center }) {
  const map = useMap();
  useEffect(() => {
    const handleResize = () => {
      map.invalidateSize();
    };
    window.addEventListener("resize", handleResize);
    const timer = setTimeout(() => map.invalidateSize(), 250);
    return () => {
      window.removeEventListener("resize", handleResize);
      clearTimeout(timer);
    };
  }, [map]);

  useEffect(() => {
    if (center) map.flyTo(center, 14);
  }, [center, map]);

  return null;
}

// ==========================================
// 2. MAIN REACT APPLICATION COMPONENT
// ==========================================
export default function EdmontonInfillDashboard() {
  // 1. Core State Hooks
  const [allSales, setAllSales] = useState([]);
  const [infillData, setInfillData] = useState([]);
  const [inventory, setInventory] = useState([]);
  const [filterLast5Years, setFilterLast5Years] = useState(false); // <-- New state
  const [evalYear, setEvalYear] = useState("");
  const [ringSize, setRingSize] = useState("");
  const [decayExponent, setDecayExponent] = useState(1.5); // Default to 1.5
  const [selectedNbhd, setSelectedNbhd] = useState("All");
  const [modelEngine, setModelEngine] = useState(
    () => new EdmontonGravityExposureModel(),
  );
  const [trainedModelData, setTrainedModelData] = useState(null);

  // 2. Selection States
  const [selectedInventoryRoll, setSelectedInventoryRoll] = useState("");
  const [predictionResult, setPredictionResult] = useState(null);
  const [mapFocusMode, setMapFocusMode] = useState("neighbourhood");

  // Loading States
  const [loadingStudy, setLoadingStudy] = useState(false);
  const [studyProgress, setStudyProgress] = useState(0);
  const [studyStatus, setStudyStatus] = useState("");
  const [loadingInventory, setLoadingInventory] = useState(false);
  const [inventoryProgress, setInventoryProgress] = useState(0);
  const [inventoryStatus, setInventoryStatus] = useState("");

  const modelFileInputRef = useRef(null);
  const inventoryFileInputRef = useRef(null);

  // Model Training Progress States
  const [trainingModel, setTrainingModel] = useState(false);
  const [trainProgress, setTrainProgress] = useState(0);
  const [trainStatus, setTrainStatus] = useState("");

  // =========================================================================
  // 3. MUST BE DECLARED HERE (BEFORE mapFilteredSales & mapFilteredInfills)
  // =========================================================================
  const selectedSubjectProperty = useMemo(() => {
    if (!selectedInventoryRoll || inventory.length === 0) return null;
    // Using String() ensures comparison works whether ROLL_NBR is a number or string
    return (
      inventory.find(
        (p) => String(p.rollNbr) === String(selectedInventoryRoll),
      ) || null
    );
  }, [inventory, selectedInventoryRoll]);

  // Automatically activate Property Buffer Focus when an inventory property is chosen
  useEffect(() => {
    if (selectedInventoryRoll) {
      setMapFocusMode("property");
    }
  }, [selectedInventoryRoll]);

  // =========================================================================
  // 4. MAP FILTERING HOOKS (Now safely accesses selectedSubjectProperty)
  // =========================================================================
  const mapFilteredSales = useMemo(() => {
    if (
      mapFocusMode === "property" &&
      selectedSubjectProperty &&
      ringSize > 0
    ) {
      return allSales.filter((s) => {
        if (s.includeFlag !== "Y") return false;
        const dist = calculateDistanceMeters(
          selectedSubjectProperty.propX,
          selectedSubjectProperty.propY,
          s.propX,
          s.propY,
        );
        return dist <= ringSize;
      });
    }

    return allSales.filter(
      (s) =>
        s.includeFlag === "Y" &&
        (selectedNbhd === "All" || s.nbhdName === selectedNbhd),
    );
  }, [allSales, selectedSubjectProperty, ringSize, mapFocusMode, selectedNbhd]);

  const mapFilteredInfills = useMemo(() => {
    if (
      mapFocusMode === "property" &&
      selectedSubjectProperty &&
      ringSize > 0
    ) {
      return infillData.filter((i) => {
        if (i.includeFlag !== "Y") return false;
        const dist = calculateDistanceMeters(
          selectedSubjectProperty.propX,
          selectedSubjectProperty.propY,
          i.infillX,
          i.infillY,
        );
        return dist <= ringSize;
      });
    }

    return infillData.filter(
      (i) =>
        i.includeFlag === "Y" &&
        (selectedNbhd === "All" || i.nbhdName === selectedNbhd),
    );
  }, [
    infillData,
    selectedSubjectProperty,
    ringSize,
    mapFocusMode,
    selectedNbhd,
  ]);

  // Helper to re-evaluate inclusion flags and spatial proximity in a single pass
  const recalculateSpatialFilters = (
    salesList,
    infillsList,
    targetYear,
    targetRing,
    use5YearFilter = filterLast5Years,
  ) => {
    if (!targetYear) {
      return { updatedSales: salesList, updatedInfills: infillsList };
    }

    // 1. Update active infills: permit year <= target year
    const updatedInfills = infillsList.map((i) => {
      const withinUpperBound = i.permitYear <= targetYear;
      const withinLowerBound = use5YearFilter
        ? i.permitYear >= targetYear - 5
        : true;
      return {
        ...i,
        includeFlag: withinUpperBound && withinLowerBound ? "Y" : "N",
      };
    });

    const activeInfills = updatedInfills.filter((i) => i.includeFlag === "Y");

    // 2. Update sales: inclusion, effective age, and ring exposure
    const updatedSales = salesList.map((s) => {
      let isIncluded =
        s.soldYear - s.actYearBuilt > 2 && targetYear >= s.soldYear; //take out infill sales built and sold within 2 years
      let units = 0;

      // Only compute Haversine distances if the sale is active and a ring size is selected
      if (isIncluded && targetRing > 0) {
        activeInfills.forEach((inf) => {
          const dist = calculateDistanceMeters(
            s.propX,
            s.propY,
            inf.infillX,
            inf.infillY,
          );
          if (dist <= 5) {
            isIncluded = false; // Less than 5 meters means the sale itself was an infill or a future infill.
          }
          if (dist <= targetRing && dist > 5) {
            units += inf.unitsAdded;
          }
        });
      }

      return {
        ...s,
        age: targetYear - s.effYearBuilt + 1,
        includeFlag: isIncluded ? "Y" : "N",
        ringSize: targetRing || 0,
        totalUnitsAdded: units,
      };
    });

    return { updatedSales, updatedInfills };
  };

  // Helper to yield control to the browser event loop for UI repainting
  const delay = (ms = 40) => new Promise((resolve) => setTimeout(resolve, ms));

  // Dynamic derived statistics (Includes Typology Breakdown)
  const stats = useMemo(() => {
    const sales = evalYear
      ? allSales.filter((s) => s.includeFlag === "Y")
      : allSales;
    const infills = evalYear
      ? infillData.filter((i) => i.includeFlag === "Y")
      : infillData;
    const soldYears = sales.map((s) => s.soldYear).filter((y) => y > 0);
    const permitYears = infills.map((i) => i.permitYear).filter((y) => y > 0);
    const compYears = infills.map((i) => i.completionYear).filter((y) => y > 0);
    const nbhds = new Set(sales.map((s) => s.nbhdName));

    // Unique typologies and project counts
    const typologyCounts = {};
    infills.forEach((i) => {
      if (i.typology) {
        typologyCounts[i.typology] = (typologyCounts[i.typology] || 0) + 1;
      }
    });
    const typologiesList = Object.keys(typologyCounts).sort();

    // ========================================================
    // NEW STATISTIC: Infill Premium Neighbourhoods
    // (Median Impacted Price > Median Non-Impacted Price)
    // ========================================================
    let infillPremiumNbhdCount = 0;
    let comparableNbhdCount = 0;

    if (ringSize > 0) {
      const grouped = {};
      sales.forEach((s) => {
        if (!grouped[s.nbhdName]) {
          grouped[s.nbhdName] = { noInfill: [], infillImpacted: [] };
        }
        if (s.totalUnitsAdded === 0) {
          grouped[s.nbhdName].noInfill.push(s.unitPriceSm);
        } else {
          grouped[s.nbhdName].infillImpacted.push(s.unitPriceSm);
        }
      });

      const median = (arr) => {
        if (!arr.length) return 0;
        const sorted = [...arr].sort((a, b) => a - b);
        return sorted[Math.floor(sorted.length / 2)];
      };

      Object.values(grouped).forEach((g) => {
        // Evaluate only neighbourhoods that have BOTH sales cohorts
        if (g.noInfill.length > 0 && g.infillImpacted.length > 0) {
          comparableNbhdCount++;
          if (median(g.infillImpacted) >= median(g.noInfill)) {
            infillPremiumNbhdCount++;
          }
        }
      });
    }

    return {
      salesCount: sales.length,
      soldYearRange: soldYears.length
        ? `${Math.min(...soldYears)} - ${Math.max(...soldYears)}`
        : "0 - 0",
      nbhdCount: nbhds.size,
      infillCount: infills.length,
      permitYearRange: permitYears.length
        ? `${Math.min(...permitYears)} - ${Math.max(...permitYears)}`
        : "0 - 0",
      compYearRange: compYears.length
        ? `${Math.min(...compYears)} - ${Math.max(...compYears)}`
        : "0 - 0",
      typologies: typologiesList,
      typologyCounts: typologyCounts,
      infillPremiumNbhdCount,
      comparableNbhdCount,
      infillPremiumPct:
        comparableNbhdCount > 0
          ? Math.round((infillPremiumNbhdCount / comparableNbhdCount) * 100)
          : 0,
    };
  }, [allSales, infillData, evalYear, ringSize]); // <-- Added ringSize to dependencies

  // 6. Longitudinal Annual Trend Data (2023 - 2026) for Selected Neighbourhood
  const neighbourhoodTrendData = useMemo(() => {
    if (selectedNbhd === "All" || allSales.length === 0) return [];

    const trendYears = [2023, 2024, 2025, 2026];
    const nbhdSales = allSales.filter((s) => s.nbhdName === selectedNbhd);
    const activeRing = ringSize > 0 ? ringSize : 50;

    const median = (arr) => {
      if (!arr.length) return null;
      const sorted = [...arr].sort((a, b) => a - b);
      return Math.round(sorted[Math.floor(sorted.length / 2)]);
    };

    return trendYears.map((year) => {
      // Sales for this neighbourhood in this specific year
      const yearSales = nbhdSales.filter(
        (s) => s.soldYear <= year && s.soldYear - s.actYearBuilt > 2,
      );

      // Active infills completed up to (year + 2)
      const activeInfills = infillData.filter((i) => i.permitYear <= year);

      const allPrices = [];
      const noInfillPrices = [];
      const impactedPrices = [];

      yearSales.forEach((s) => {
        let units = 0;
        if (activeRing > 0) {
          activeInfills.forEach((inf) => {
            const d = calculateDistanceMeters(
              s.propX,
              s.propY,
              inf.infillX,
              inf.infillY,
            );
            if (d <= activeRing) {
              units += inf.unitsAdded;
            }
          });
        }
        const unadjUnitPrice =
          s.price > 0 && s.flrAreaSm > 0 ? s.price / s.flrAreaSm : 0; //calculate unadjusted unit price
        allPrices.push(unadjUnitPrice);
        if (units === 0) {
          noInfillPrices.push(unadjUnitPrice);
        } else {
          impactedPrices.push(unadjUnitPrice);
        }
      });

      return {
        year: String(year),
        "Median Price (All)": median(allPrices),
        "Count (All)": allPrices.length,

        "Median Price (0 Infills)": median(noInfillPrices),
        "Count (0 Infills)": noInfillPrices.length,

        "Median Price (>0 Infills)": median(impactedPrices),
        "Count (>0 Infills)": impactedPrices.length,
      };
    });
  }, [allSales, infillData, selectedNbhd, ringSize]);

  // Dynamic Y-Axis scale for the 3-series Line Chart
  // Domain: [min - 200, max + 200] divided into 5 equal intervals (6 tick marks)
  const lineChartYAxis = useMemo(() => {
    if (!neighbourhoodTrendData || neighbourhoodTrendData.length === 0) {
      return {
        domain: [3000, 6000],
        ticks: [3000, 3600, 4200, 4800, 5400, 6000],
      };
    }

    // Extract all valid numeric prices across all 3 series
    const values = [];
    neighbourhoodTrendData.forEach((row) => {
      [
        "Median Price (All)",
        "Median Price (0 Infills)",
        "Median Price (>0 Infills)",
      ].forEach((key) => {
        const val = row[key];
        if (typeof val === "number" && !isNaN(val) && val > 0) {
          values.push(val);
        }
      });
    });

    if (values.length === 0) {
      return {
        domain: [3000, 6000],
        ticks: [3000, 3600, 4200, 4800, 5400, 6000],
      };
    }

    const rawMin = Math.min(...values);
    const rawMax = Math.max(...values);

    // Dynamic Domain: Min - 200 to Max + 200
    const minY = Math.floor(rawMin - 200);
    const maxY = Math.ceil(rawMax + 200);

    // 5 equal intervals = 6 boundary tick marks
    const step = (maxY - minY) / 5;
    const ticks = Array.from({ length: 6 }, (_, i) =>
      Math.round(minY + i * step),
    );

    return {
      domain: [minY, maxY],
      ticks: ticks,
    };
  }, [neighbourhoodTrendData]);

  // Automatically activate Property Buffer Focus whenever a roll number is picked
  useEffect(() => {
    if (selectedInventoryRoll) {
      setMapFocusMode("property");
    }
  }, [selectedInventoryRoll]);

  // Chart Aggregations
  const chartData = useMemo(() => {
    let sales = evalYear
      ? allSales.filter((s) => s.includeFlag === "Y")
      : allSales;
    if (selectedNbhd !== "All")
      sales = sales.filter((s) => s.nbhdName === selectedNbhd);

    const grouped = {};
    sales.forEach((s) => {
      if (!grouped[s.nbhdName])
        grouped[s.nbhdName] = {
          nbhd: s.nbhdName,
          all: [],
          noInfill: [],
          infillImpacted: [],
        };
      grouped[s.nbhdName].all.push(s.unitPriceSm);
      if (s.totalUnitsAdded === 0)
        grouped[s.nbhdName].noInfill.push(s.unitPriceSm);
      else grouped[s.nbhdName].infillImpacted.push(s.unitPriceSm);
    });

    const median = (arr) => {
      if (!arr.length) return 0;
      const sorted = [...arr].sort((a, b) => a - b);
      return sorted[Math.floor(sorted.length / 2)];
    };

    return Object.values(grouped).map((g) => ({
      nbhd: g.nbhd,
      "Median Price (All)": median(g.all),
      "Count (All)": g.all.length,
      "Median Price (0 Infills)": median(g.noInfill),
      "Count (0 Infills)": g.noInfill.length,
      "Median Price (>0 Infills)": median(g.infillImpacted),
      "Count (>0 Infills)": g.infillImpacted.length,
    }));
  }, [allSales, evalYear, selectedNbhd]);

  // --- HANDLERS ---
  const handleLoadStudyData = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      setLoadingStudy(true);
      setStudyProgress(15);
      setStudyStatus("Reading binary buffer from file...");
      await delay();

      const buffer = await file.arrayBuffer();
      setStudyProgress(45);
      setStudyStatus("Parsing workbook sheets...");
      await delay();

      const wb = read(buffer, { type: "array" });
      setStudyProgress(70);
      setStudyStatus("Extracting AllSales & InfillData...");
      await delay();

      const salesSheet = wb.Sheets["AllSales"] || wb.Sheets[wb.SheetNames[0]];
      const infillsSheet =
        wb.Sheets["InfillData"] || wb.Sheets[wb.SheetNames[1]];
      const rawSales = utils.sheet_to_json(salesSheet);
      const rawInfills = utils.sheet_to_json(infillsSheet);

      setStudyProgress(90);
      setStudyStatus("Initializing spatial attributes & model arrays...");
      await delay();

      const engine = new EdmontonGravityExposureModel();
      const processed = engine.load_study_data(rawSales, rawInfills);

      setAllSales(processed.allSales);
      setInfillData(processed.infillData);
      setModelEngine(engine);

      setStudyProgress(100);
      setStudyStatus("Complete!");
      await delay(350); // Brief hold so the user sees 100%
    } catch (err) {
      alert("Error loading study data: " + err.message);
    } finally {
      setLoadingStudy(false);
      setStudyProgress(0);
      setStudyStatus("");
      e.target.value = ""; // Reset input to allow re-uploading the same file
    }
  };

  // Handle Evaluation Year Change
  const handleEvalYearChange = (e) => {
    const year = parseInt(e.target.value, 10);
    if (year === 2026) {
      setEvalYear(year);
    } else {
      alert("No future sales for this analysis.");
      return;
    }

    // Synchronously re-evaluate both infills and sales using the existing ringSize
    const { updatedSales, updatedInfills } = recalculateSpatialFilters(
      allSales,
      infillData,
      year,
      ringSize,
      filterLast5Years,
    );

    setAllSales(updatedSales);
    setInfillData(updatedInfills);
  };

  // Handle 5-Year Permit Radio Toggle
  const handle5YearFilterChange = (e) => {
    const isEnabled = e.target.value === "5yr";
    setFilterLast5Years(isEnabled);

    // Immediately re-run spatial filters if an evaluation year is already active
    if (evalYear) {
      const { updatedSales, updatedInfills } = recalculateSpatialFilters(
        allSales,
        infillData,
        evalYear,
        ringSize,
        isEnabled,
      );
      setAllSales(updatedSales);
      setInfillData(updatedInfills);
    }
  };

  // Handle Ring Size Change
  const handleRingSizeChange = (e) => {
    if (!evalYear) {
      alert("Please select Evaluation Year first.");
      return;
    }

    const ring = parseInt(e.target.value, 10);
    setRingSize(ring);

    // Synchronously re-evaluate both infills and sales using the current evalYear
    const { updatedSales, updatedInfills } = recalculateSpatialFilters(
      allSales,
      infillData,
      evalYear,
      ring,
      filterLast5Years,
    );

    setAllSales(updatedSales);
    setInfillData(updatedInfills);
  };

  // const handleTrainModel = () => {
  //   const engine = new EdmontonGravityExposureModel(evalYear, ringSize);
  //   const trained = engine.fit(allSales, infillData);
  //   if (trained) {
  //     setTrainedModelData(trained);
  //     setModelEngine(engine);
  //   }
  // };

  const handleTrainModel = async () => {
    if (!evalYear) {
      alert("Please select Evaluation Year first.");
      return;
    }
    if (!ringSize) {
      alert("Please select Ring Size first.");
      return;
    }

    try {
      setTrainingModel(true);
      setTrainProgress(15);
      setTrainStatus(`Filtering sales & infills up to ${evalYear}...`);
      await delay(50);

      // 1. Ensure latest spatial calculations and filters are synchronized
      const { updatedSales, updatedInfills } = recalculateSpatialFilters(
        allSales,
        infillData,
        evalYear,
        ringSize,
        filterLast5Years,
      );
      setAllSales(updatedSales);
      setInfillData(updatedInfills);

      setTrainProgress(35);
      setTrainStatus(
        `Computing distance-decay gravity exposure vectors (${ringSize}m)...`,
      );
      await delay(50);

      // PASS DECAY EXPONENT HERE:
      const engine = new EdmontonGravityExposureModel(
        evalYear,
        ringSize,
        decayExponent,
      );

      setTrainProgress(60);
      setTrainStatus(
        "Absorbing neighbourhood fixed effects (MAP algorithm)...",
      );
      await delay(50);

      setTrainProgress(85);
      setTrainStatus(
        "Estimating OLS hedonic parameters via matrix inversion...",
      );
      await delay(50);

      const trained = engine.fit(updatedSales, updatedInfills);
      if (!trained) {
        throw new Error("Matrix estimation failed. Try a larger ring size.");
      }

      // Update state with fresh object references to guarantee component re-render
      setModelEngine(engine);
      setTrainedModelData({ ...trained });

      // Automatically re-evaluate selected property with the new coefficients
      if (selectedInventoryRoll && inventory.length > 0) {
        try {
          const updatedPred = engine.predict(
            selectedInventoryRoll,
            inventory,
            updatedInfills,
          );
          setPredictionResult({ ...updatedPred });
        } catch (err) {
          console.warn("Could not auto-refresh prediction:", err);
        }
      }

      setTrainProgress(100);
      setTrainStatus(`Model trained for ${evalYear} (${ringSize}m)!`);
      await delay(350);
    } catch (err) {
      alert("Model Training Error: " + err.message);
    } finally {
      setTrainingModel(false);
      setTrainProgress(0);
      setTrainStatus("");
    }
  };

  // Feature 1: Load Existing Model from JSON
  const handleLoadModelJSON = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const parsed = JSON.parse(evt.target.result);
          const engine = new EdmontonGravityExposureModel();
          const loaded = engine.load_model(parsed);
          setModelEngine(engine);
          setTrainedModelData(loaded);
          if (loaded.eval_year) setEvalYear(loaded.eval_year);
          if (loaded.ring_size) setRingSize(loaded.ring_size);
          // SYNC DECAY EXPONENT FROM LOADED JSON
          if (loaded.decay_exponent) setDecayExponent(loaded.decay_exponent);
        } catch (err) {
          alert("Invalid model JSON file: " + err.message);
        }
      };
      reader.readAsText(file);
    }
  };

  // Feature 2: Load Property Inventory Spreadsheet
  const handleLoadInventorySpreadsheet = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      setLoadingInventory(true);
      setInventoryProgress(20);
      setInventoryStatus("Reading inventory file buffer...");
      await delay();

      const buffer = await file.arrayBuffer();
      setInventoryProgress(50);
      setInventoryStatus("Parsing property inventory sheet...");
      await delay();

      const wb = read(buffer, { type: "array" });
      const sheet = wb.Sheets["RandomProperty"] || wb.Sheets[wb.SheetNames[0]];
      const rawInv = utils.sheet_to_json(sheet);

      setInventoryProgress(85);
      setInventoryStatus(
        "Structuring property coordinates & characteristics...",
      );
      await delay();

      const parsedInv = modelEngine.load_property_inventory(
        rawInv,
        evalYear || 2026,
      );
      setInventory(parsedInv);
      if (parsedInv.length > 0) setSelectedInventoryRoll(parsedInv[0].rollNbr);

      setInventoryProgress(100);
      setInventoryStatus("Inventory ready!");
      await delay(350);
    } catch (err) {
      alert("Error loading inventory file: " + err.message);
    } finally {
      setLoadingInventory(false);
      setInventoryProgress(0);
      setInventoryStatus("");
      e.target.value = "";
    }
  };

  // Feature 2: Execute Prediction
  const handlePredictInventory = () => {
    if (!trainedModelData || !selectedInventoryRoll) return;
    try {
      const result = modelEngine.predict(
        selectedInventoryRoll,
        inventory,
        infillData,
      );
      setPredictionResult(result);
    } catch (err) {
      alert("Prediction error: " + err.message);
    }
  };

  const exportCSV = (data, filename) => {
    const ws = utils.json_to_sheet(data);
    const csv = utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
  };

  return (
    <Container
      maxWidth="xl"
      sx={{
        py: { xs: 1.5, sm: 2, md: 3 },
        px: { xs: 1, sm: 2, md: 3 },
      }}
    >
      <Typography
        variant="h4"
        fontWeight="bold"
        sx={{
          fontSize: { xs: "1.4rem", sm: "1.8rem", md: "2.1rem" },
          mb: { xs: 1.5, md: 2.5 },
        }}
      >
        Infills Impact Analysis - Gravity Exposure Model (COE)
      </Typography>
      {/* ========================================================
          TWO-TRACK CSS GRID CONTAINER
          Left Track: min 500px | Right Track: 1fr (100% of remaining width)
         ======================================================== */}
      <Box
        sx={{
          display: "grid",
          gridTemplateColumns: {
            xs: "1fr",
            md: "500px 1fr", // Exactly 500px on medium screens, rest to col 2
            lg: "minmax(500px, 28%) 1fr", // Min 500px or 28% on desktops, rest to col 2
            xl: "minmax(500px, 24%) 1fr", // Min 500px or 24% on wide screens, rest to col 2
          },
          gap: { xs: 2, md: 2.5 },
          width: "100%",
          alignItems: "stretch",
        }}
      >
        {/* ========================================================
            ROW 1 - LEFT: User Input
            (Row 1, Col 1)
           ======================================================== */}
        <Paper
          elevation={2}
          sx={{
            p: { xs: 2, lg: 2.5 },
            display: "flex",
            flexDirection: "column",
            gap: 1.75,
            borderRadius: 2,
            minWidth: 0,
          }}
        >
          <Typography variant="h6" fontWeight="bold">
            User Input
          </Typography>

          {/* Study Data Loader Button + Progress */}
          <Button
            variant="contained"
            component="label"
            disabled={loadingStudy}
            fullWidth
          >
            {loadingStudy
              ? "Loading Study Data..."
              : "Load Data (InfillStudyData.xlsx)"}
            <input
              type="file"
              hidden
              accept=".xlsx"
              onChange={handleLoadStudyData}
            />
          </Button>

          {loadingStudy && (
            <Box sx={{ width: "100%", mt: -0.5 }}>
              <Box
                sx={{
                  display: "flex",
                  justifyContent: "space-between",
                  mb: 0.5,
                }}
              >
                <Typography variant="caption" color="textSecondary">
                  {studyStatus}
                </Typography>
                <Typography variant="caption" fontWeight="bold">
                  {studyProgress}%
                </Typography>
              </Box>
              <LinearProgress
                variant="determinate"
                value={studyProgress}
                sx={{ height: 6, borderRadius: 3 }}
              />
            </Box>
          )}

          <FormControl disabled={allSales.length === 0} fullWidth size="small">
            <InputLabel>Evaluation Year</InputLabel>
            <Select
              value={evalYear}
              label="Evaluation Year"
              onChange={handleEvalYearChange}
            >
              {[2026, 2027].map((y) => (
                <MenuItem key={y} value={y}>
                  {y}
                </MenuItem>
              ))}
            </Select>
          </FormControl>

          {/* --- NEW: 5-YEAR INFILL PERMIT FILTER RADIO GROUP --- */}
          <FormControl
            disabled={!evalYear || allSales.length === 0}
            component="fieldset"
            sx={{
              px: 1.25,
              py: 0.5,
              border: "1px solid #e0e0e0",
              borderRadius: 1,
              bgcolor: filterLast5Years ? "#f3e5f5" : "transparent",
            }}
          >
            <FormLabel
              component="legend"
              sx={{ fontSize: "0.75rem", fontWeight: "bold" }}
            >
              Infill Permit Window
            </FormLabel>
            <RadioGroup
              row
              value={filterLast5Years ? "5yr" : "all"}
              onChange={handle5YearFilterChange}
            >
              <FormControlLabel
                value="all"
                control={<Radio size="small" />}
                label={
                  <Typography variant="caption">All Prior Permits</Typography>
                }
              />
              <FormControlLabel
                value="5yr"
                control={<Radio size="small" color="secondary" />}
                label={
                  <Typography
                    variant="caption"
                    fontWeight={filterLast5Years ? "bold" : "normal"}
                  >
                    Previous 5 Years{" "}
                    {evalYear ? `(${evalYear - 5}–${evalYear})` : ""}
                  </Typography>
                }
              />
            </RadioGroup>
          </FormControl>

          <FormControl
            disabled={!evalYear || allSales.length === 0}
            fullWidth
            size="small"
          >
            <InputLabel>Ring Size (Meters)</InputLabel>
            <Select
              value={ringSize}
              label="Ring Size (Meters)"
              onChange={handleRingSizeChange}
            >
              {[15, 30, 50, 100, 250].map((r) => (
                <MenuItem key={r} value={r}>
                  {r}m
                </MenuItem>
              ))}
            </Select>
          </FormControl>

          {/* --- NEW: DECAY EXPONENT CHOICE BOX --- */}
          <FormControl disabled={!ringSize} fullWidth size="small">
            <InputLabel>Decay Exponent (α)</InputLabel>
            <Select
              value={decayExponent}
              label="Decay Exponent (α)"
              onChange={(e) => setDecayExponent(parseFloat(e.target.value))}
            >
              <MenuItem value={1.5}>1.5 (Standard Distance Decay)</MenuItem>
              <MenuItem value={2.0}>2.0 (Inverse Square Law)</MenuItem>
            </Select>
          </FormControl>

          {/* <Button variant="contained" color="secondary" disabled={allSales.length === 0} onClick={handleTrainModel} fullWidth>
            Train Exposure Model
          </Button> */}

          {/* --- MODEL TRAINING BUTTON + PROGRESS BAR --- */}
          <Button
            variant="contained"
            color="secondary"
            disabled={!ringSize || !evalYear || trainingModel}
            onClick={handleTrainModel}
            fullWidth
          >
            {trainingModel
              ? "Training Exposure Model..."
              : "Train Exposure Model"}
          </Button>

          {trainingModel && (
            <Box sx={{ width: "100%", mt: -0.5 }}>
              <Box
                sx={{
                  display: "flex",
                  justifyContent: "space-between",
                  mb: 0.5,
                }}
              >
                <Typography variant="caption" color="textSecondary">
                  {trainStatus}
                </Typography>
                <Typography variant="caption" fontWeight="bold">
                  {trainProgress}%
                </Typography>
              </Box>
              <LinearProgress
                variant="determinate"
                value={trainProgress}
                color="secondary"
                sx={{ height: 6, borderRadius: 3 }}
              />
            </Box>
          )}

          <Divider sx={{ my: 0.25 }} />

          {/* Model Management */}
          <Typography variant="caption" fontWeight="bold" color="textSecondary">
            MODEL STORAGE & IMPORT
          </Typography>
          <Grid container spacing={1}>
            <Grid size={{ xs: 6 }}>
              <Button
                variant="outlined"
                component="label"
                size="small"
                fullWidth
                sx={{ fontSize: "0.75rem", py: 0.8 }}
              >
                Load JSON
                <input
                  type="file"
                  hidden
                  accept=".json"
                  ref={modelFileInputRef}
                  onChange={handleLoadModelJSON}
                />
              </Button>
            </Grid>
            <Grid size={{ xs: 6 }}>
              <Button
                variant="outlined"
                size="small"
                disabled={!trainedModelData}
                fullWidth
                sx={{ fontSize: "0.75rem", py: 0.8 }}
                onClick={() => {
                  const blob = new Blob(
                    [JSON.stringify(trainedModelData, null, 2)],
                    { type: "application/json" },
                  );
                  const link = document.createElement("a");
                  link.href = URL.createObjectURL(blob);
                  link.download = `EdmontonGravityModel_${evalYear || "Saved"}.json`;
                  link.click();
                }}
              >
                Export JSON
              </Button>
            </Grid>
          </Grid>

          <Divider sx={{ my: 0.25 }} />

          {/* Property Inventory Prediction */}
          <Typography variant="caption" fontWeight="bold" color="textSecondary">
            PROPERTY INVENTORY PREDICTION
          </Typography>
          <Button
            variant="contained"
            color="info"
            size="small"
            disabled={!trainedModelData || loadingInventory}
            component="label"
            fullWidth
          >
            {loadingInventory ? "Loading Inventory..." : "Load Inventory Sheet"}
            <input
              type="file"
              hidden
              accept=".xlsx,.csv"
              ref={inventoryFileInputRef}
              onChange={handleLoadInventorySpreadsheet}
            />
          </Button>

          {loadingInventory && (
            <Box sx={{ width: "100%", mt: -0.5 }}>
              <Box
                sx={{
                  display: "flex",
                  justifyContent: "space-between",
                  mb: 0.5,
                }}
              >
                <Typography variant="caption" color="textSecondary">
                  {inventoryStatus}
                </Typography>
                <Typography variant="caption" fontWeight="bold">
                  {inventoryProgress}%
                </Typography>
              </Box>
              <LinearProgress
                variant="determinate"
                value={inventoryProgress}
                color="info"
                sx={{ height: 6, borderRadius: 3 }}
              />
            </Box>
          )}

          <FormControl disabled={inventory.length === 0} fullWidth size="small">
            <InputLabel>Select Property Roll</InputLabel>
            <Select
              value={selectedInventoryRoll}
              label="Select Property Roll"
              onChange={(e) => setSelectedInventoryRoll(e.target.value)}
            >
              {inventory.map((p) => (
                <MenuItem key={p.rollNbr} value={p.rollNbr}>
                  Roll: {p.rollNbr} ({p.nbhdName})
                </MenuItem>
              ))}
            </Select>
          </FormControl>

          <Button
            variant="contained"
            color="success"
            disabled={!selectedInventoryRoll || !trainedModelData}
            onClick={handlePredictInventory}
            fullWidth
          >
            Predict Property Value
          </Button>

          {inventory.length > 0 && (
            <Button
              size="small"
              onClick={() => exportCSV(inventory, "Inventory_Valuations.csv")}
            >
              Export Evaluated CSV
            </Button>
          )}
        </Paper>

        {/* ========================================================
            ROW 1 - RIGHT: Statistics Visualization
            (Row 1, Col 2 - Takes over remaining width beside User Input)
           ======================================================== */}
        <Paper
          elevation={2}
          sx={{
            p: { xs: 2, lg: 2.5 },
            display: "flex",
            flexDirection: "column",
            borderRadius: 2,
            minWidth: 0,
          }}
        >
          <Typography variant="h6" fontWeight="bold" gutterBottom>
            Statistics Visualization
          </Typography>

          {/* Quick Stat Badges: Auto-scales across all screen sizes */}
          <Box
            sx={{
              display: "grid",
              gridTemplateColumns: {
                xs: "repeat(2, 1fr)",
                sm: "repeat(3, 1fr)",
                md: "repeat(4, 1fr)",
                xl: "repeat(7, 1fr)",
              },
              gap: 1.25,
              mb: 1.5,
            }}
          >
            {/* 1. Sales Count */}
            <Alert severity="info" icon={false} sx={{ py: 0.5, px: 1.2 }}>
              <Typography variant="caption" display="block">
                Sales Count
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {stats.salesCount.toLocaleString()}
              </Typography>
            </Alert>

            {/* 2. Sold Year Range */}
            <Alert severity="info" icon={false} sx={{ py: 0.5, px: 1.2 }}>
              <Typography variant="caption" display="block">
                Sold Year Range
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {stats.soldYearRange}
              </Typography>
            </Alert>

            {/* 3. Total Neighbourhoods */}
            <Alert severity="info" icon={false} sx={{ py: 0.5, px: 1.2 }}>
              <Typography variant="caption" display="block">
                Neighbourhoods
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {stats.nbhdCount}
              </Typography>
            </Alert>

            {/* 4. Infill Projects */}
            <Alert severity="success" icon={false} sx={{ py: 0.5, px: 1.2 }}>
              <Typography variant="caption" display="block">
                Infill Count
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {stats.infillCount.toLocaleString()}
              </Typography>
            </Alert>

            {/* 5. Infill Permit Year Range */}
            <Alert severity="success" icon={false} sx={{ py: 0.5, px: 1.2 }}>
              <Typography variant="caption" display="block">
                Infill Permit Year Range
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {stats.permitYearRange}
              </Typography>
            </Alert>

            {/* 6. Typologies Count */}
            <Alert
              severity="secondary"
              icon={false}
              sx={{ py: 0.5, px: 1.2, bgcolor: "#f3e5f5", color: "#4a148c" }}
            >
              <Typography variant="caption" display="block">
                Typologies Count
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {stats.typologies.length} Types
              </Typography>
            </Alert>

            {/* 7. NEW: Infill Premium Nbhds (Impacted > Non-Impacted) */}
            <Alert
              severity="warning"
              icon={false}
              sx={{
                py: 0.5,
                px: 1.2,
                bgcolor: ringSize ? "#fff8e1" : undefined,
                color: ringSize ? "#b78103" : undefined,
              }}
            >
              <Typography variant="caption" display="block" noWrap>
                Impacted &gt; 0-Infill
              </Typography>
              <Typography variant="subtitle2" fontWeight="bold">
                {ringSize > 0 ? (
                  <span>
                    {stats.infillPremiumNbhdCount} / {stats.comparableNbhdCount}{" "}
                    <span style={{ fontSize: "0.72rem", fontWeight: "normal" }}>
                      ({stats.infillPremiumPct}%)
                    </span>
                  </span>
                ) : (
                  <span
                    style={{
                      fontSize: "0.75rem",
                      fontStyle: "italic",
                      color: "#888",
                    }}
                  >
                    Select Ring Size
                  </span>
                )}
              </Typography>
            </Alert>
          </Box>

          {/* Full List of Typologies (Styled Chips with Counts) */}
          <Box
            sx={{
              display: "flex",
              alignItems: "center",
              gap: 1,
              flexWrap: "wrap",
              mb: 2,
              px: 0.5,
            }}
          >
            <Typography
              variant="caption"
              fontWeight="bold"
              color="textSecondary"
            >
              Infill Typologies ({stats.typologies.length}):
            </Typography>
            {stats.typologies.length > 0 ? (
              stats.typologies.map((typ) => (
                <Chip
                  key={typ}
                  size="small"
                  label={`${typ} (${stats.typologyCounts[typ].toLocaleString()})`}
                  color="secondary"
                  variant="outlined"
                  sx={{ fontWeight: "medium", fontSize: "0.75rem" }}
                />
              ))
            ) : (
              <Typography
                variant="caption"
                color="textSecondary"
                fontStyle="italic"
              >
                Load data to view infill types and counts
              </Typography>
            )}
          </Box>

          {/* Neighbourhood Filter */}
          <FormControl fullWidth size="small" sx={{ mb: 1.5 }}>
            <InputLabel>Filter Neighbourhood</InputLabel>
            <Select
              value={selectedNbhd}
              label="Filter Neighbourhood"
              onChange={(e) => setSelectedNbhd(e.target.value)}
            >
              <MenuItem value="All">All Mature Neighbourhoods</MenuItem>
              {Array.from(new Set(allSales.map((s) => s.nbhdName)))
                .filter(Boolean)
                .sort((a, b) => a.localeCompare(b))
                .map((n) => (
                  <MenuItem key={n} value={n}>
                    {n}
                  </MenuItem>
                ))}
            </Select>
          </FormControl>

          {/* 1. Bar Chart Container */}
          {evalYear > 0 && (
            <Box
              sx={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                mb: 0.5,
              }}
            >
              <Typography variant="subtitle1" fontWeight="bold">
                Median Time Adjusted Unit Price Comparison
              </Typography>
              <Chip
                label={`Evaulation Year: ${evalYear || 2026}`}
                size="small"
                color="secondary"
                variant="outlined"
              />
            </Box>
          )}
          <Box
            sx={{
              flexGrow: 1,
              minHeight: { xs: 260, sm: 300, md: 350 },
              width: "100%",
              minWidth: 0,
            }}
          >
            <ResponsiveContainer width="100%" height="100%">
              <BarChart
                data={chartData}
                margin={{ top: 10, right: 10, left: -10, bottom: 0 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="nbhd"
                  tick={{ fontSize: 11 }}
                  interval={0}
                  angle={-25}
                  textAnchor="end"
                  height={55}
                />
                <YAxis
                  tickFormatter={(val) => `$${val / 1000}k`}
                  tick={{ fontSize: 11 }}
                />
                <Tooltip
                  formatter={(value, name, item) => {
                    const row = item.payload; // Accesses the entire chartData row
                    // Match the count property to the series being hovered
                    let count = 0;
                    if (name === "All Sales in Nbhd") {
                      count = row["Count (All)"] ?? row.countAll ?? 0;
                    } else if (name === "No Infills (0 Units)") {
                      count =
                        row["Count (0 Infills)"] ?? row.countNoInfill ?? 0;
                    } else if (name === "Infill Impacted (>0 Units)") {
                      count = row["Count (>0 Infills)"] ?? row.countInfill ?? 0;
                    }

                    // Returns: [Formatted Value, Series Name]
                    return [
                      `$${Number(value).toLocaleString()} (${count} sale${count === 1 ? "" : "s"})`,
                      name,
                    ];
                  }}
                />
                <Legend position="top" offset={10} />
                {ringSize ? (
                  <>
                    <Bar
                      dataKey="Median Price (0 Infills)"
                      fill="#8884d8"
                      name="No Infills (0 Units)"
                    />
                    <Bar
                      dataKey="Median Price (All)"
                      fill="#82ca9d"
                      name="All Sales in Nbhd"
                    />
                    <Bar
                      dataKey="Median Price (>0 Infills)"
                      fill="#ffc658"
                      name="Infill Impacted (>0 Units)"
                    />
                  </>
                ) : (
                  <Bar
                    dataKey="Median Price (All)"
                    fill="#82ca9d"
                    name="All Sales in Nbhd"
                  />
                )}
              </BarChart>
            </ResponsiveContainer>
          </Box>

          {/* ==================================================================================
              2. THREE-SERIES TREND LINE CHART (2023 - 2026
              Only rendered when a specific neighbourhood is selected and a ring size is selected
             ================================================================================== */}
          {selectedNbhd !== "All" && ringSize > 0 && (
            <Box
              sx={{
                mt: 2.5,
                pt: 2,
                borderTop: "1px solid #e0e0e0",
                minWidth: 0,
              }}
            >
              <Box
                sx={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  mb: 0.5,
                }}
              >
                <Typography variant="subtitle1" fontWeight="bold">
                  Unadjusted Historical Sales Trend (2023 – 2026: {selectedNbhd}
                  )
                </Typography>
                <Chip
                  label={`Ring Buffer: ${ringSize || 50}m`}
                  size="small"
                  color="secondary"
                  variant="outlined"
                />
              </Box>

              <Typography
                variant="caption"
                color="textSecondary"
                sx={{ display: "block", mb: 1.5 }}
              >
                Median single-family unit price ($/m²). Number of sales (n) is
                labeled at each point.
              </Typography>

              <Box sx={{ height: 280, width: "100%", minWidth: 0 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart
                    data={neighbourhoodTrendData}
                    margin={{ top: 20, right: 25, left: -10, bottom: 5 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis
                      dataKey="year"
                      tick={{ fontSize: 12, fontWeight: "bold" }}
                    />
                    {/* <YAxis
                      // domain={[3000, 6000]}
                      // ticks={[3000, 3500, 4000, 4500, 5000, 5500, 6000]}
                      tickFormatter={(val) => `$${val / 1000}k`}
                      tick={{ fontSize: 11 }}
                      // allowDataOverflow={true}
                    /> */}
                    <YAxis
                      domain={lineChartYAxis.domain}
                      ticks={lineChartYAxis.ticks}
                      tickFormatter={(val) =>
                        `$${Math.round(val).toLocaleString()}`
                      }
                      tick={{ fontSize: 11 }}
                      allowDataOverflow={true}
                    />

                    {/* Interactive Tooltip showing all 3 cohorts for that year */}
                    <Tooltip
                      formatter={(value, name, item) => {
                        const row = item?.payload || {};
                        const activeKey = item?.dataKey || "";
                        let count = 0;
                        if (activeKey === "Median Price (0 Infills)")
                          count = row["Count (0 Infills)"] ?? 0;
                        else if (activeKey === "Median Price (>0 Infills)")
                          count = row["Count (>0 Infills)"] ?? 0;
                        else count = row["Count (All)"] ?? 0;

                        return value
                          ? [
                              `$${Number(value).toLocaleString()} (${count} sales)`,
                              name,
                            ]
                          : ["No Data", name];
                      }}
                    />

                    <Legend position="top" offset={10} />

                    {/* Series 1: Non-Impacted Sales (Purple) */}
                    <Line
                      type="monotone"
                      dataKey="Median Price (0 Infills)"
                      name="No Infills (0 Units)"
                      stroke="#8884d8"
                      strokeWidth={2.5}
                      connectNulls
                      dot={{ r: 5, fill: "#8884d8" }}
                      activeDot={{ r: 7 }}
                    >
                      <LabelList
                        dataKey="Count (0 Infills)"
                        position="top"
                        offset={8}
                        formatter={(val) => (val > 0 ? `n=${val}` : "")}
                        style={{
                          fontSize: "10px",
                          fontWeight: "bold",
                          fill: "#5c6bc0",
                        }}
                      />
                    </Line>

                    {/* Series 2: All Sales (Green) */}
                    <Line
                      type="monotone"
                      dataKey="Median Price (All)"
                      name="All Sales in Nbhd"
                      stroke="#82ca9d"
                      strokeWidth={3}
                      connectNulls
                      dot={{ r: 5, fill: "#82ca9d" }}
                      activeDot={{ r: 7 }}
                    >
                      <LabelList
                        dataKey="Count (All)"
                        position="bottom"
                        offset={9}
                        formatter={(val) => (val > 0 ? `n=${val}` : "")}
                        style={{
                          fontSize: "10px",
                          fontWeight: "bold",
                          fill: "#2e7d32",
                        }}
                      />
                    </Line>

                    {/* Series 3: Infill Impacted Sales (Orange / Amber) */}
                    <Line
                      type="monotone"
                      dataKey="Median Price (>0 Infills)"
                      name="Infill Impacted (>0 Units)"
                      stroke="#ff9800"
                      strokeWidth={2.5}
                      connectNulls
                      dot={{ r: 5, fill: "#ff9800" }}
                      activeDot={{ r: 7 }}
                    >
                      <LabelList
                        dataKey="Count (>0 Infills)"
                        position="top"
                        offset={8}
                        formatter={(val) => (val > 0 ? `n=${val}` : "")}
                        style={{
                          fontSize: "10px",
                          fontWeight: "bold",
                          fill: "#e65100",
                        }}
                      />
                    </Line>
                  </LineChart>
                </ResponsiveContainer>
              </Box>
            </Box>
          )}
        </Paper>

        {/* ========================================================
            ROW 2 - LEFT: Sales/Infills Map (MIN WIDTH 500PX)
           ======================================================== */}
        <Paper
          elevation={2}
          sx={{
            p: { xs: 2, lg: 2.5 },
            display: "flex",
            flexDirection: "column",
            borderRadius: 2,
            minWidth: 0,
            position: "relative",
            zIndex: 0,
            overflow: "hidden",
          }}
        >
          {/* Header & View Mode Switcher */}
          <Box
            sx={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              mb: 0.5,
            }}
          >
            <Typography variant="h6" fontWeight="bold">
              Sales/Infills Map
            </Typography>

            {selectedSubjectProperty && (
              <Box sx={{ display: "flex", gap: 1 }}>
                <Chip
                  label={
                    mapFocusMode === "property"
                      ? `Ring: ${ringSize}m Buffer`
                      : "Nbhd View"
                  }
                  size="small"
                  color={mapFocusMode === "property" ? "success" : "default"}
                  onClick={() =>
                    setMapFocusMode((m) =>
                      m === "property" ? "neighbourhood" : "property",
                    )
                  }
                  sx={{ cursor: "pointer", fontWeight: "bold" }}
                />
              </Box>
            )}
          </Box>

          <Typography variant="caption" color="textSecondary" sx={{ mb: 1 }}>
            {mapFocusMode === "property" && selectedSubjectProperty ? (
              <span>
                Isolated to <strong>{ringSize}m buffer</strong> of Roll #
                {selectedSubjectProperty.rollNbr}:{" "}
                <strong>{mapFilteredSales.length} sales</strong>,{" "}
                <strong>{mapFilteredInfills.length} infills</strong>
              </span>
            ) : (
              "Red: Infill | Blue: Sales | Green: Subject Inventory Property"
            )}
          </Typography>

          {/* Leaflet Map */}
          <Box
            sx={{
              flexGrow: 1,
              position: "relative",
              borderRadius: 1.5,
              overflow: "hidden",
              minHeight: 500,
              width: "100%",
              zIndex: 0,
            }}
          >
            <MapContainer
              center={[53.5461, -113.4938]}
              zoom={12}
              style={{ height: "100%", width: "100%" }}
            >
              <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />

              <MapViewController
                selectedNbhd={selectedNbhd}
                filteredSales={mapFilteredSales}
                filteredInfills={mapFilteredInfills}
                subjectProp={selectedSubjectProperty}
                ringSize={ringSize}
                focusMode={mapFocusMode}
              />

              <LayersControl position="topright">
                {/* BASE LAYER */}
                <LayersControl.BaseLayer checked name="OpenStreetMap Standard">
                  <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                </LayersControl.BaseLayer>

                {/* OVERLAY 1: SALES */}
                <LayersControl.Overlay
                  checked
                  name="Single-Family Sales (Blue)"
                >
                  <LayerGroup>
                    {mapFilteredSales.map((s, idx) => (
                      <CircleMarker
                        key={`sale-${s.id}-${s.rollNbr || ""}-${idx}`}
                        center={[s.propX, s.propY]}
                        radius={4}
                        pathOptions={{
                          color: "#0d47a1",
                          fillColor: "#2196f3",
                          fillOpacity: 0.7,
                        }}
                      >
                        {/* HOVER TOOLTIP */}
                        <LeafletTooltip
                          direction="top"
                          offset={[0, -4]}
                          opacity={0.95}
                        >
                          <div style={{ fontSize: "12px", lineHeight: 1.4 }}>
                            <strong>Sold Property (Roll: {s.rollNbr})</strong>
                            <br />
                            Sale {s.saleDate ? "Date" : "Year"}:{" "}
                            <strong>{s.saleDate || s.soldYear}</strong>
                            <br />
                            Time-Adjusted Price:{" "}
                            <strong>
                              ${Math.round(s.unitPriceSm).toLocaleString()}/m²
                            </strong>
                            {s.unadjUnitPriceSm != null && (
                              <>
                                <br />
                                Unadjusted Price:{" "}
                                <strong>
                                  $
                                  {Math.round(
                                    s.unadjUnitPriceSm,
                                  ).toLocaleString()}
                                  /m²
                                </strong>
                              </>
                            )}
                          </div>
                        </LeafletTooltip>
                        {/* CLICK POPUP */}
                        <Popup>
                          Roll: {s.rollNbr}
                          <br />
                          Sale {s.saleDate ? "Date" : "Year"}:{" "}
                          {s.saleDate || s.soldYear}
                          <br />
                          Adjusted Price: $
                          {Math.round(s.unitPriceSm).toLocaleString()}/m²
                          {s.unadjUnitPriceSm != null && (
                            <>
                              <br />
                              formatExcelDate)( Unadjusted Price: $
                              {Math.round(s.unadjUnitPriceSm).toLocaleString()}
                              /m²
                            </>
                          )}
                        </Popup>
                      </CircleMarker>
                    ))}
                  </LayerGroup>
                </LayersControl.Overlay>

                {/* OVERLAY 2: INFILLS */}
                <LayersControl.Overlay checked name="Infill Projects (Red)">
                  <LayerGroup>
                    {mapFilteredInfills.map((i, idx) => (
                      <CircleMarker
                        key={`inf-${i.infillId || idx}`}
                        center={[i.infillX, i.infillY]}
                        radius={5}
                        pathOptions={{
                          color: "#b71c1c",
                          fillColor: "#f44336",
                          fillOpacity: 0.9,
                        }}
                      >
                        {/* HOVER TOOLTIP */}
                        <LeafletTooltip
                          direction="top"
                          offset={[0, -5]}
                          opacity={0.95}
                        >
                          <div style={{ fontSize: "12px", lineHeight: 1.4 }}>
                            <strong>Infill #{i.infillId}</strong>
                            <br />
                            Type: <strong>{i.typology}</strong>
                            <br />
                            Units Added: <strong>+{i.unitsAdded}</strong>
                            <br />
                            Permit {i.permitDate ? "Date" : "Year"}:{" "}
                            <strong>
                              {i.permitDate || i.permitYear || "N/A"}
                            </strong>
                          </div>
                        </LeafletTooltip>

                        {/* CLICK POPUP */}
                        <Popup>
                          <strong>Infill #{i.infillId}</strong>
                          <br />
                          Type: {i.typology}
                          <br />
                          Units: +{i.unitsAdded}
                          <br />
                          Permit {i.permitDate ? "Date" : "Year"}:{" "}
                          {i.permitDate || i.permitYear || "N/A"}
                        </Popup>
                      </CircleMarker>
                    ))}
                  </LayerGroup>
                </LayersControl.Overlay>

                {/* OVERLAY 3: SUBJECT PROPERTY */}
                {selectedSubjectProperty && (
                  <LayersControl.Overlay
                    checked
                    name="Subject Property & Buffer (Green)"
                  >
                    <LayerGroup>
                      <CircleMarker
                        center={[
                          selectedSubjectProperty.propX,
                          selectedSubjectProperty.propY,
                        ]}
                        radius={8}
                        pathOptions={{
                          color: "#1b5e20",
                          fillColor: "#00e676",
                          fillOpacity: 1.0,
                        }}
                      >
                        <Popup>Roll: {selectedSubjectProperty.rollNbr}</Popup>
                      </CircleMarker>

                      {ringSize > 0 && (
                        <Circle
                          center={[
                            selectedSubjectProperty.propX,
                            selectedSubjectProperty.propY,
                          ]}
                          radius={ringSize}
                          pathOptions={{
                            color: "#2e7d32",
                            fillColor: "#a5d6a7",
                            fillOpacity: 0.25,
                            dashArray: "4, 4",
                          }}
                        />
                      )}
                    </LayerGroup>
                  </LayersControl.Overlay>
                )}
              </LayersControl>
            </MapContainer>
          </Box>
        </Paper>

        {/* ========================================================
            ROW 2 - RIGHT: Model Summary
            (Row 2, Col 2 - Automatically takes over ALL remaining width beside Map Pane)
           ======================================================== */}
        <Paper
          elevation={2}
          sx={{
            p: { xs: 2, lg: 2.5 },
            display: "flex",
            flexDirection: "column",
            borderRadius: 2,
            minWidth: 0,
            overflow: "hidden",
          }}
        >
          {/* <Typography variant="h6" fontWeight="bold">Model Summary</Typography>
           */}
          {/* Display the decay exponent from the trained model data or the default value */}
          <Box
            sx={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <Typography variant="h6" fontWeight="bold">
              Model Summary
            </Typography>
            {trainedModelData && (
              <Box sx={{ display: "flex", gap: 1 }}>
                <Chip
                  label={`Year: ${evalYear}`}
                  size="small"
                  variant="outlined"
                />
                <Chip
                  label={`Ring: ${ringSize}m`}
                  size="small"
                  variant="outlined"
                />
                <Chip
                  label={`Decay α: ${trainedModelData.decay_exponent || decayExponent}`}
                  size="small"
                  color="secondary"
                  variant="outlined"
                />
                {trainedModelData.calibrationStats && (
                  <>
                    {/* RAW INITIAL ASR & COD */}
                    {/* <Chip
                      label={`Raw ASR: ${trainedModelData.calibrationStats.rawMedianASR} (COD: ${trainedModelData.calibrationStats.rawCOD}%)`}
                      size="small"
                      variant="outlined"
                    /> */}

                    {/* FINAL CALIBRATED ASR & COD */}
                    <Chip
                      label={`Final ASR: ${trainedModelData.calibrationStats.finalMedianASR.toFixed(3)} (COD: ${trainedModelData.calibrationStats.finalCOD}%)`}
                      size="small"
                      color="success"
                      variant="filled"
                    />

                    {/* MODIFIED NEIGHBOURHOOD COUNT */}
                    <Chip
                      label={`NBHD Modifiers (|diff|>= ${NBHDMODIFIERTHRESHOLD}): ${trainedModelData.modifierStats.adjusted} / ${trainedModelData.modifierStats.total}`}
                      size="small"
                      color={
                        trainedModelData.modifierStats.adjusted > 0
                          ? "warning"
                          : "default"
                      }
                      variant="outlined"
                    />
                  </>
                )}
              </Box>
            )}
          </Box>
          <Typography variant="caption" color="textSecondary" sx={{ mb: 1.5 }}>
            Econometric specification, regression parameters, and out-of-sample
            property valuations
          </Typography>

          {trainedModelData ? (
            <Box
              sx={{
                display: "flex",
                flexDirection: "column",
                gap: 2,
                flexGrow: 1,
                minWidth: 0,
              }}
            >
              {/* Mathematical Formula Box (wordBreak + overflowWrap prevents horizontal blowout) */}
              <Box
                sx={{
                  bgcolor: "#263238",
                  color: "#eceff1",
                  p: 1.5,
                  borderRadius: 1.5,
                  fontFamily: "monospace",
                  overflowX: "auto",
                  minWidth: 0,
                }}
              >
                <Typography
                  variant="body2"
                  sx={{
                    whiteSpace: "pre-wrap",
                    wordBreak: "break-all",
                    fontSize: { xs: "0.75rem", md: "0.82rem" },
                  }}
                >
                  {trainedModelData.formula}
                </Typography>
              </Box>
              {trainedModelData.sampleAverages && (
                <Box
                  sx={{
                    display: "grid",
                    gridTemplateColumns: { xs: "1fr", md: "1fr 1fr" },
                    gap: 1.5,
                    my: 1,
                  }}
                >
                  {/* Card 1: Estimation Sample Averages */}
                  <Paper variant="outlined" sx={{ p: 1.5, bgcolor: "#f9fbe7" }}>
                    <Typography
                      variant="caption"
                      fontWeight="bold"
                      color="textSecondary"
                      display="block"
                    >
                      ESTIMATION SAMPLE CENTROID (N ={" "}
                      {trainedModelData.sampleAverages.salesCount})
                    </Typography>
                    <Typography variant="body2" sx={{ mt: 0.5 }}>
                      Typical House:{" "}
                      <strong>
                        {trainedModelData.sampleAverages.typicalLivingArea} m²
                      </strong>{" "}
                      (Lot: {trainedModelData.sampleAverages.typicalLotSize} m²)
                    </Typography>
                    <Typography variant="body2">
                      Typical Quality:{" "}
                      <strong>
                        {
                          trainedModelData.sampleAverages.typicalQuality.split(
                            ".",
                          )[0]
                        }
                      </strong>{" "}
                      | Typical Condition:{" "}
                      <strong>
                        {
                          trainedModelData.sampleAverages.typicalCondition.split(
                            ".",
                          )[0]
                        }
                      </strong>
                    </Typography>
                    <Typography variant="body2">
                      Typical Age:{" "}
                      <strong>
                        {trainedModelData.sampleAverages.typicalAge} yrs
                      </strong>{" "}
                      | Median Price:{" "}
                      <strong>
                        $
                        {trainedModelData.sampleAverages.typicalUnitPrice.toLocaleString()}
                        /m²
                      </strong>
                    </Typography>
                  </Paper>

                  {/* Card 2: Standard Benchmark Baseline */}
                  {/* <Paper variant="outlined" sx={{ p: 1.5, bgcolor: "#e8f5e9" }}>
                    <Typography variant="caption" fontWeight="bold" color="success.dark" display="block">
                      BENCHMARK PROPERTY BASELINE (0 INFILL EXPOSURE)
                    </Typography>
                    <Typography variant="body2" sx={{ mt: 0.5 }}>
                      Specs: <strong>{trainedModelData.benchmark.specs.flrAreaSm} m²</strong>, {trainedModelData.benchmark.specs.lotSizeSm} m² lot, {trainedModelData.benchmark.specs.age} yrs old
                    </Typography>
                    <Typography variant="body2">
                      Baseline Unit Price: <strong>${trainedModelData.benchmark.benchmarkUnitPrice.toLocaleString()}/m²</strong> (${trainedModelData.benchmark.benchmarkTotalValue.toLocaleString()})
                    </Typography>
                  </Paper> */}
                </Box>
              )}

              <Grid container spacing={2} sx={{ flexGrow: 1 }}>
                {/* Regression Parameter Table */}
                <Grid size={{ xs: 12, lg: 7 }}>
                  <TableContainer
                    component={Paper}
                    variant="outlined"
                    sx={{ maxHeight: 500, minWidth: 0 }}
                  >
                    <Table size="small" stickyHeader>
                      <TableHead>
                        <TableRow>
                          <TableCell>
                            <strong>Variable</strong>
                          </TableCell>
                          <TableCell align="right">
                            <strong>Coefficient ($\beta$)</strong>
                          </TableCell>
                          <TableCell align="right">
                            <strong>% Impact</strong>
                          </TableCell>
                        </TableRow>
                      </TableHead>
                      <TableBody>
                        {modelEngine.get_full_regression_table().map((row) => (
                          <TableRow key={row.variable}>
                            <TableCell
                              sx={{ fontFamily: "monospace", fontSize: 12 }}
                            >
                              {row.variable}
                            </TableCell>
                            <TableCell align="right">
                              {row.coefficient.toFixed(5)}
                            </TableCell>
                            <TableCell align="right">
                              <Chip
                                size="small"
                                label={`${row.pct_impact >= 0 ? "+" : ""}${row.pct_impact.toFixed(2)}%`}
                                color={
                                  row.pct_impact >= 0 ? "success" : "error"
                                }
                                variant="outlined"
                              />
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </TableContainer>
                </Grid>

                {/* Valuation Result Card */}
                <Grid size={{ xs: 12, lg: 5 }}>
                  {predictionResult ? (
                    <Card
                      variant="outlined"
                      sx={{
                        bgcolor: "#fbfdf8",
                        borderColor: "#c0ca33",
                        height: "100%",
                      }}
                    >
                      <CardContent sx={{ p: 2 }}>
                        <Typography
                          variant="subtitle1"
                          fontWeight="bold"
                          color="primary"
                          gutterBottom
                        >
                          Valuation: Roll #{predictionResult.rollNbr}
                        </Typography>
                        <Typography variant="body2">
                          Neighbourhood:{" "}
                          <strong>{predictionResult.nbhdName}</strong>
                        </Typography>
                        <Typography variant="body2">
                          Neighbourhood Intercept:{" "}
                          <strong>
                            {predictionResult.nbhdIntercept.toFixed(4)}
                          </strong>
                        </Typography>
                        <Typography variant="body2">
                          Floor Area:{" "}
                          <strong>
                            {predictionResult.flrAreaSm.toFixed(2)} m²
                          </strong>{" "}
                          | Age: <strong>{predictionResult.age} yrs</strong>
                        </Typography>
                        {/* --- TOTAL ASSESSMENT (CITY OF EDMONTON) --- */}
                        <Box sx={{ mb: 1 }}>
                          <Typography
                            variant="caption"
                            color="textSecondary"
                            display="block"
                          >
                            Assessed Value:
                          </Typography>
                          <Typography
                            variant="h6"
                            fontWeight="bold"
                            color="textPrimary"
                          >
                            {predictionResult.assessedValue > 0
                              ? `$${Math.round(predictionResult.assessedValue).toLocaleString()}`
                              : "N/A"}
                          </Typography>
                        </Box>
                        {/* NBHD Modifier Status in Valuation Card */}
                        {predictionResult.nbhdModifier && (
                          <Box
                            sx={{
                              display: "flex",
                              justifyContent: "space-between",
                              alignItems: "center",
                              mt: 1,
                              mb: 1,
                            }}
                          >
                            <Typography variant="caption" color="textSecondary">
                              NBHD Modifier ({predictionResult.nbhdName}):
                            </Typography>
                            <Chip
                              size="small"
                              label={
                                predictionResult.nbhdModifier !== 1.0
                                  ? `Modifier: ×${predictionResult.nbhdModifier}`
                                  : "1.000"
                              }
                              color={
                                predictionResult.nbhdModifier !== 1.0
                                  ? "warning"
                                  : "default"
                              }
                              variant="outlined"
                              sx={{
                                height: 22,
                                fontSize: "0.72rem",
                                fontWeight: "bold",
                              }}
                            />
                          </Box>
                        )}

                        <Divider sx={{ my: 1.2 }} />

                        <Typography
                          variant="caption"
                          color="textSecondary"
                          display="block"
                        >
                          Estimated Unit Price:
                        </Typography>
                        <Typography
                          variant="h6"
                          color="success.main"
                          fontWeight="bold"
                        >
                          $
                          {Math.round(
                            predictionResult.predUnitPrice,
                          ).toLocaleString()}{" "}
                          / m²
                        </Typography>

                        <Typography
                          variant="caption"
                          color="textSecondary"
                          display="block"
                          sx={{ mt: 0.5 }}
                        >
                          Estimated Total Value (With Infill Exposure):
                        </Typography>
                        <Typography variant="h5" fontWeight="bold">
                          $
                          {Math.round(
                            predictionResult.predTotalValue,
                          ).toLocaleString()}
                        </Typography>

                        {/* --- NEW: COUNTERFACTUAL COMPARISON (0 GRAVITY EXPOSURE) --- */}
                        {predictionResult.noInfillTotalValue !== undefined && (
                          <Box
                            sx={{
                              mt: 1.25,
                              p: 1.25,
                              borderRadius: 1.5,
                              bgcolor: "#f1f8e9",
                              border: "1px dashed #aed581",
                            }}
                          >
                            <Typography
                              variant="caption"
                              fontWeight="bold"
                              color="textSecondary"
                              display="block"
                            >
                              Baseline Value (0 Infill Exposure):
                            </Typography>
                            <Box
                              sx={{
                                display: "flex",
                                justifyContent: "space-between",
                                alignItems: "baseline",
                                mt: 0.25,
                              }}
                            >
                              <Typography variant="subtitle1" fontWeight="bold">
                                $
                                {Math.round(
                                  predictionResult.noInfillTotalValue,
                                ).toLocaleString()}
                              </Typography>
                              <Typography
                                variant="caption"
                                color="textSecondary"
                              >
                                ($
                                {Math.round(
                                  predictionResult.noInfillUnitPrice,
                                ).toLocaleString()}
                                /m²)
                              </Typography>
                            </Box>

                            <Box
                              sx={{
                                display: "flex",
                                justifyContent: "space-between",
                                alignItems: "center",
                                mt: 0.75,
                              }}
                            >
                              <Typography
                                variant="caption"
                                color="textSecondary"
                              >
                                Net Infill Spillover Impact:
                              </Typography>
                              {Object.values(predictionResult.exposures).some(
                                (v) => v > 0,
                              ) ? (
                                <Chip
                                  size="small"
                                  label={`${
                                    predictionResult.infillImpactDollar >= 0
                                      ? "+"
                                      : "-"
                                  }$${Math.abs(
                                    Math.round(
                                      predictionResult.infillImpactDollar,
                                    ),
                                  ).toLocaleString()} (${
                                    predictionResult.infillImpactPct >= 0
                                      ? "+"
                                      : ""
                                  }${predictionResult.infillImpactPct.toFixed(2)}%)`}
                                  color={
                                    predictionResult.infillImpactDollar >= 0
                                      ? "success"
                                      : "error"
                                  }
                                  variant="outlined"
                                  sx={{
                                    height: 22,
                                    fontSize: "0.72rem",
                                    fontWeight: "bold",
                                  }}
                                />
                              ) : (
                                <Chip
                                  size="small"
                                  label="$0 (0.00% — No Nearby Infills)"
                                  variant="outlined"
                                  sx={{ height: 22, fontSize: "0.72rem" }}
                                />
                              )}
                            </Box>
                          </Box>
                        )}

                        <Divider sx={{ my: 1.2 }} />

                        <Typography
                          variant="caption"
                          fontWeight="bold"
                          display="block"
                        >
                          Gravity Exposures:
                        </Typography>
                        <Grid container spacing={0.5} sx={{ mt: 0.2 }}>
                          {Object.entries(predictionResult.exposures).map(
                            ([typ, score]) => (
                              <Grid size={{ xs: 6 }} key={typ}>
                                <Typography
                                  variant="caption"
                                  color="textSecondary"
                                  display="block"
                                >
                                  {typ}: <strong>{score.toFixed(5)}</strong>
                                </Typography>
                              </Grid>
                            ),
                          )}
                        </Grid>
                      </CardContent>
                    </Card>
                  ) : (
                    <Alert
                      severity="info"
                      sx={{
                        height: "100%",
                        display: "flex",
                        alignItems: "center",
                      }}
                    >
                      Load an Inventory Sheet, select a property, and click{" "}
                      <strong>Predict Property Value</strong>.
                    </Alert>
                  )}
                </Grid>
              </Grid>
            </Box>
          ) : (
            <Box
              sx={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                flexGrow: 1,
                minHeight: 280,
              }}
            >
              <Typography variant="body2" color="textSecondary">
                No model trained or loaded yet. Complete the steps in the User
                Input to view model formulas and regression estimates.
              </Typography>
            </Box>
          )}
        </Paper>
      </Box>{" "}
    </Container>
  );
}
