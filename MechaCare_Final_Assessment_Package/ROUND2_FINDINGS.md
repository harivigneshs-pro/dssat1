# ROUND 2 Findings

## 1. Dataset overview
- sensor.csv: 31,011 rows, 5 columns
- machines.csv: 60 rows, 4 columns
- usage.csv: 2,340 rows, 5 columns
- alarms.csv: 949 rows, 4 columns
- maintenance.csv: 106 rows, 5 columns

## 2. Data quality findings
- Duplicate sensors: 1 exact duplicate and 1 duplicate `(machine, timestamp)` key after cleaning.
- Missing sensor values: 12 vibration, 12 temperature, 11 current.
- Alarm labels contained mixed case and spacing; these were normalized to canonical labels.
- Maintenance downtime had 2 missing values; action had 3 missing values; these were retained with `NaN`/`unknown` handling.

## 3. Business question answers
### Q1. Normal and abnormal sensor ranges
- Vibrations were summarized with min, median, mean, and IQR outlier bounds.
- The raw dataset showed outlier counts above the IQR rule: vibration 720, temperature 671, current 344.
- These flagged values were not deleted; they were retained for review.

### Q2. Sensor readings before maintenance
- For each machine, sensor records were aligned to the next service event using `days_to_next_service`.
- The analysis was restricted to the realistic window of 0-21 days before service.
- Readings were grouped into 0-3, 4-7, 8-14, and 15-21 day windows.

### Q3. Which alarm types occur most frequently?
- Normalized labels were used to count alarm events.
- The dominant alarm categories were `general_warning`, `high_vibration`, and `high_temperature`, after canonical mapping.

### Q4. Load vs downtime
- Machine-level aggregation was used to avoid repeating downtime values across every weekly usage row.
- The machine-level correlation between mean load and total known downtime was weak and did not support a causal claim.

### Q5. Which machine types require more servicing?
- Machine types were compared using service events per machine, total known downtime, and repeat-service machines.
- The strongest supported comparison is the per-machine rate, not total counts alone.

### Q6. Repeated failure patterns
- Repeated maintenance actions, repeated alarms, high outlier counts, and recurring service events were used to build a transparent inspection priority list.
- The final ranking is a descriptive shortlist and not a validated predictive maintenance model.

## 4. File outputs
- Cleaned datasets saved under `cleaned/`
- Summary and output CSVs saved under `outputs/`
- Charts saved under `plots/`
- Final written summary saved as `ROUND2_FINDINGS.md`

## 5. Limitations
- The analysis is descriptive and retrospective.
- Missing downtime and action data create partial maintenance histories.
- No causal or predictive claim is justified from correlation alone.
