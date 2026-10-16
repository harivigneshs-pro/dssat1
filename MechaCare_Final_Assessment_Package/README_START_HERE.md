# MechaCare Machine Maintenance Insight — Final Assessment Package

## Important data note
The included CSV files are **synthetic** and were generated for the workflow because the assessment PDF describes the schema but does not include the underlying operational CSV data. Do not describe these as real MechaCare records. If the organizers provide official CSV files, replace the files in `raw/` and rerun the notebook and workbook.

## Files
- `raw/sensor.csv`, `raw/machines.csv`, `raw/usage.csv`, `raw/alarms.csv`, `raw/maintenance.csv`
- `cleaned/sensor_machine_analysis.csv`
- `cleaned/machine_weekly_analysis.csv`
- `cleaned/machine_maintenance_summary.csv`
- `outputs/machine_priority_ranking.csv`, alarm/type/before-service summaries
- `MechaCare_EDA_Feature_Engineering.ipynb`
- `MechaCare_Maintenance_Analysis.xlsx`
- `MechaCare_Business_Insight_Pitch.pptx`

## Start-to-finish workflow
1. Extract the ZIP into one folder.
2. Open the notebook in Jupyter or VS Code with Python.
3. Ensure `pandas`, `numpy`, `matplotlib`, `seaborn`, and `openpyxl` are installed.
4. Run cells from top to bottom. The notebook expects `raw/` beside the notebook. If running from this package root, that is already true.
5. Inspect output tables and plots; do not copy numeric claims from the synthetic-data pitch if using different official files.
6. Open the Excel workbook. Review `Summary`, `weekly_analysis`, `alarm_frequency`, `type_comparison`, `before_service`, and `priority_ranking`.
7. Update the pitch with actual values and explain evidence → business meaning → recommendation.

## Data quality decisions
- Duplicate sensor machine/timestamp rows: keep the first record for the prepared analysis; preserve raw data.
- Extreme sensor readings: flag them and retain originals; create clipped analysis columns for robust summaries. Confirm true equipment limits with engineers before removing any values.
- Alarm labels: standardize spelling/case/spacing; unknown labels remain visible.
- Missing service downtime: treat as unknown, not zero at event level.
- Missing weekly alarm aggregation: blank means no matched event row / unknown coverage, not guaranteed zero.
- Join strategy: sensor + machines at reading level; aggregate sensor and alarms to machine-week before joining usage; maintenance remains event/machine-level.

## Short pitch structure
1. Business question and objective.
2. Data quality and method.
3. Three evidence-backed findings.
4. Business impact / risk.
5. Prioritized recommendations and limitations.

## Suggested quiz revision
- Data types, mean/median/std, IQR and outlier handling.
- Missingness and duplicates; normalization and encoding.
- Pandas `read_csv`, `info`, `isna`, `duplicated`, `groupby`, `merge`.
- NumPy arrays and vectorized operations.
- Matplotlib/Seaborn bar, histogram, box, line and scatter plots.
- Correlation is association, not causation; avoid data leakage in feature engineering.
