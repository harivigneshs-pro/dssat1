from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference, ScatterChart
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

root = Path(r"C:\Users\Hari Vignesh S\Downloads\MechaCare_Final_Assessment_Package1\MechaCare_Final_Assessment_Package")
output_path = root / "MechaCare_Round3_Final.xlsx"

sensor = pd.read_csv(root / "cleaned" / "sensor_clean.csv")
maintenance_clean = pd.read_csv(root / "cleaned" / "maintenance_clean.csv")
machine_summary = pd.read_csv(root / "outputs" / "machine_level_summary.csv")
machine_type = pd.read_csv(root / "outputs" / "machine_type_comparison.csv")
alarm_counts = pd.read_csv(root / "outputs" / "alarm_frequency.csv")
pre_maintenance = pd.read_csv(root / "outputs" / "pre_maintenance_sensor_summary.csv")
load_downtime = pd.read_csv(root / "outputs" / "load_downtime_analysis.csv")
priority = pd.read_csv(root / "outputs" / "machine_priority_ranking.csv")

# Business metrics from the project data
machine_count = len(machine_summary)
service_event_total = int(maintenance_clean['service_event_id'].nunique())
problem_total = int(alarm_counts['count'].sum())
mean_load = float(load_downtime['mean_load'].mean())
mean_downtime = float(load_downtime['total_known_downtime'].mean())
load_corr = float(load_downtime[['mean_load', 'total_known_downtime']].corr().iloc[0, 1])
q1_outlier = {
    'vibration': int(((sensor['vibration'] < sensor['vibration'].quantile(0.25) - 1.5 * (sensor['vibration'].quantile(0.75) - sensor['vibration'].quantile(0.25))) |
                      (sensor['vibration'] > sensor['vibration'].quantile(0.75) + 1.5 * (sensor['vibration'].quantile(0.75) - sensor['vibration'].quantile(0.25)))).sum()),
    'temperature': int(((sensor['temperature'] < sensor['temperature'].quantile(0.25) - 1.5 * (sensor['temperature'].quantile(0.75) - sensor['temperature'].quantile(0.25))) |
                       (sensor['temperature'] > sensor['temperature'].quantile(0.75) + 1.5 * (sensor['temperature'].quantile(0.75) - sensor['temperature'].quantile(0.25)))).sum()),
    'current': int(((sensor['current'] < sensor['current'].quantile(0.25) - 1.5 * (sensor['current'].quantile(0.75) - sensor['current'].quantile(0.25))) |
                    (sensor['current'] > sensor['current'].quantile(0.75) + 1.5 * (sensor['current'].quantile(0.75) - sensor['current'].quantile(0.25)))).sum())
}

def set_header(ws, headers, row=1):
    fill = PatternFill("solid", fgColor="1F4E78")
    for col_idx, value in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=value)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )


def auto_filter_and_styles(ws, start_row=1, end_row=None, end_col=None):
    if end_row is None:
        end_row = ws.max_row
    if end_col is None:
        end_col = ws.max_column
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(end_col)}{end_row}"
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=1, max_col=end_col):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )
    for col in ws.columns:
        max_len = 0
        for cell in col:
            val = cell.value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 24)


def set_dataframe(ws, df, start_row=1, title=None):
    if title:
        ws.title = title
    rows = df.fillna("").values.tolist()
    for r_idx, row in enumerate(rows, start=start_row):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    if start_row == 1:
        set_header(ws, list(df.columns), row=1)
        data_start = 2
    else:
        data_start = start_row + 1
    data_end_row = data_start + len(df.index) - 1
    end_col = len(df.columns)
    for row in range(data_start, data_end_row + 1):
        for col in range(1, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = Alignment(vertical="center")
    auto_filter_and_styles(ws, start_row=1, end_row=data_end_row, end_col=end_col)


wb = Workbook()
ws = wb.active
ws.title = "Management_Summary"
ws.merge_cells("A1:H1")
ws["A1"] = "MechaCare Machine Maintenance Insight — Round 3 Executive Summary"
ws["A1"].font = Font(size=16, bold=True, color="1F1F1F")
ws["A1"].fill = PatternFill("solid", fgColor="D9EAF7")
ws["A1"].alignment = Alignment(horizontal="center")

summary_rows = [
    ["Metric", "Value"],
    ["Total machines", "=COUNTA('Machine_Summary'!A:A)-1"],
    ["Total maintenance events", "=SUM('Machine_Type_Analysis'!C:C)"],
    ["Total alarm events", "=SUM('Alarm_Analysis'!B2:B10)"],
    ["Mean machine load", round(mean_load, 2)],
    ["Mean downtime per machine", round(mean_downtime, 2)],
    ["Load vs downtime correlation", round(load_corr, 4)],
    ["IQR outlier count (vibration)", q1_outlier['vibration']],
    ["IQR outlier count (temperature)", q1_outlier['temperature']],
    ["IQR outlier count (current)", q1_outlier['current']],
]
for row_idx, row in enumerate(summary_rows, start=2):
    for col_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        if row_idx == 2:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
        if row_idx > 2:
            cell.alignment = Alignment(horizontal="left" if col_idx == 1 else "center")
            if col_idx == 2 and isinstance(value, (int, float)):
                cell.number_format = '#,##0.00'

ws["A12"] = "Business Question / Finding"
ws["B12"] = "Evidence"
for i, (q, evidence) in enumerate([
    ("Q1. Normal and abnormal sensor ranges", "IQR-based outlier counts: vibration 720, temperature 671, current 344. These were retained with flags; not deleted."),
    ("Q2. Readings before maintenance", "Sensor readings were aligned to each machine's next service event using a 0–21 day window. The peak readings occur in the 0–7 day window before service."),
    ("Q3. Most frequent alarm types", "Normalized alarm counts: general_warning 636, high_vibration 218, high_temperature 64, overcurrent 26."),
    ("Q4. Load vs downtime", "Machine-level correlation between average load and total downtime is 0.0409, which is weak and not causal."),
    ("Q5. Machine types with more servicing", "Compressor and CNC show the highest average services per machine, while Pump has the highest machine count and significant downtime."),
    ("Q6. Priority shortlist", "Top ranked machines by the priority score are M053, M004, M042, M027, and M011; this is descriptive triage, not a production predictive model."),
], start=13):
    ws.cell(row=i+13, column=1, value=q)
    ws.cell(row=i+13, column=2, value=evidence)
    ws.cell(row=i+13, column=1).alignment = Alignment(wrap_text=True)
    ws.cell(row=i+13, column=2).alignment = Alignment(wrap_text=True)

ws["A39"] = "Key recommendation"
ws["B39"] = "Use the cleaned sensor, alarm, and maintenance datasets as the evidence base for inspection scheduling; do not infer causality from the weak load/downtime correlation alone."
ws["A39"].font = Font(bold=True)
ws["A39"].alignment = Alignment(wrap_text=True)
ws["B39"].alignment = Alignment(wrap_text=True)
ws["A41"] = "Limitation"
ws["B41"] = "Some maintenance action and downtime records are incomplete; the analysis is descriptive and retrospective, not a validated predictive maintenance model."
ws["A41"].font = Font(bold=True)
ws["A41"].alignment = Alignment(wrap_text=True)
ws["B41"].alignment = Alignment(wrap_text=True)

for col in ['A', 'B']:
    ws.column_dimensions[col].width = 32
for row in range(1, 50):
    ws.row_dimensions[row].height = 22
ws.freeze_panes = 'A2'

# Sheet 2: Cleaned sensor data
ws_sensor = wb.create_sheet("Cleaned_Sensor_Data")
rows = sensor.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_sensor.cell(row=r_idx, column=c_idx, value=value)
set_header(ws_sensor, list(sensor.columns), row=1)
auto_filter_and_styles(ws_sensor, start_row=1, end_row=ws_sensor.max_row, end_col=ws_sensor.max_column)

# Sheet 3: Machine summary
ws_machine = wb.create_sheet("Machine_Summary")
priority_score = priority[["machine", "warning_score"]].rename(columns={"warning_score": "priority_score"})
machine_sheet = machine_summary.merge(priority_score, on="machine", how="left")
rows = machine_sheet.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_machine.cell(row=r_idx, column=c_idx, value=value)
header = list(machine_sheet.columns)
set_header(ws_machine, header, row=1)
auto_filter_and_styles(ws_machine, start_row=1, end_row=ws_machine.max_row, end_col=ws_machine.max_column)

# Sheet 4: Machine Type Analysis
ws_type = wb.create_sheet("Machine_Type_Analysis")
machine_type_work = machine_type.copy()
machine_type_work['avg_downtime_per_service'] = machine_type_work['total_known_downtime'] / machine_type_work['total_service_events']
rows = machine_type_work.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_type.cell(row=r_idx, column=c_idx, value=value)
set_header(ws_type, list(machine_type_work.columns), row=1)
auto_filter_and_styles(ws_type, start_row=1, end_row=ws_type.max_row, end_col=ws_type.max_column)

# Sheet 5: Alarm Analysis
ws_alarm = wb.create_sheet("Alarm_Analysis")
rows = alarm_counts.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_alarm.cell(row=r_idx, column=c_idx, value=value)
set_header(ws_alarm, list(alarm_counts.columns), row=1)
auto_filter_and_styles(ws_alarm, start_row=1, end_row=ws_alarm.max_row, end_col=ws_alarm.max_column)
chart_alarm = BarChart()
chart_alarm.title = "Normalized Alarm Counts"
chart_alarm.y_axis.title = "Count"
chart_alarm.x_axis.title = "Alarm Type"
chart_alarm.style = 10
chart_alarm.height = 7
chart_alarm.width = 12
cats = Reference(ws_alarm, min_col=1, min_row=2, max_row=ws_alarm.max_row)
data = Reference(ws_alarm, min_col=2, min_row=1, max_row=ws_alarm.max_row)
chart_alarm.add_data(data, titles_from_data=True)
chart_alarm.set_categories(cats)
ws_alarm.add_chart(chart_alarm, "E3")

# Sheet 6: Maintenance Analysis
ws_maint = wb.create_sheet("Maintenance_Analysis")
maint_summary = (maintenance_clean.groupby('action', as_index=False)
                 .agg(event_count=('service_event_id', 'count'), total_downtime=('downtime', 'sum'), avg_downtime=('downtime', 'mean'), machines_affected=('machine', 'nunique'))
                 .sort_values('event_count', ascending=False)
                 .reset_index(drop=True))
rows = maint_summary.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_maint.cell(row=r_idx, column=c_idx, value=value)
set_header(ws_maint, list(maint_summary.columns), row=1)
auto_filter_and_styles(ws_maint, start_row=1, end_row=ws_maint.max_row, end_col=ws_maint.max_column)
chart_maint = BarChart()
chart_maint.title = "Maintenance Events by Action"
chart_maint.y_axis.title = "Event Count"
chart_maint.x_axis.title = "Action"
chart_maint.style = 11
chart_maint.height = 7
chart_maint.width = 12
cats_maint = Reference(ws_maint, min_col=1, min_row=2, max_row=ws_maint.max_row)
data_maint = Reference(ws_maint, min_col=2, min_row=1, max_row=ws_maint.max_row)
chart_maint.add_data(data_maint, titles_from_data=True)
chart_maint.set_categories(cats_maint)
ws_maint.add_chart(chart_maint, "F3")

# Sheet 7: Pre Maintenance Analysis
ws_pre = wb.create_sheet("Pre_Maintenance_Analysis")
rows = pre_maintenance.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_pre.cell(row=r_idx, column=c_idx, value=value)
set_header(ws_pre, list(pre_maintenance.columns), row=1)
auto_filter_and_styles(ws_pre, start_row=1, end_row=ws_pre.max_row, end_col=ws_pre.max_column)

# Sheet 8: Load vs Downtime
ws_load = wb.create_sheet("Load_vs_Downtime")
load_downtime_sheet = load_downtime.copy()
rows = load_downtime_sheet.fillna("").values.tolist()
for r_idx, row in enumerate(rows, start=1):
    for c_idx, value in enumerate(row, start=1):
        ws_load.cell(row=r_idx, column=c_idx, value=value)
set_header(ws_load, list(load_downtime_sheet.columns), row=1)
auto_filter_and_styles(ws_load, start_row=1, end_row=ws_load.max_row, end_col=ws_load.max_column)
chart_scatter = ScatterChart()
chart_scatter.title = "Machine Load vs Total Known Downtime"
chart_scatter.x_axis.title = "Average Load"
chart_scatter.y_axis.title = "Total Downtime (hours)"
chart_scatter.style = 13
chart_scatter.height = 7
chart_scatter.width = 12
x_vals = Reference(ws_load, min_col=2, min_row=2, max_row=ws_load.max_row)
y_vals = Reference(ws_load, min_col=3, min_row=1, max_row=ws_load.max_row)
chart_scatter.add_data(y_vals, titles_from_data=True)
chart_scatter.set_categories(x_vals)
ws_load.add_chart(chart_scatter, "F3")

# Sheet 9: Sensor Analysis
ws_sensor_analysis = wb.create_sheet("Sensor_Analysis")
ws_sensor_analysis["A1"] = "Metric"
ws_sensor_analysis["B1"] = "Count"
ws_sensor_analysis["C1"] = "Mean"
ws_sensor_analysis["D1"] = "Std Dev"
ws_sensor_analysis["E1"] = "Min"
ws_sensor_analysis["F1"] = "Max"
for row_idx, metric in enumerate(['vibration', 'temperature', 'current'], start=2):
    s = sensor[metric].dropna()
    desc = s.describe()
    ws_sensor_analysis.cell(row=row_idx, column=1, value=metric)
    ws_sensor_analysis.cell(row=row_idx, column=2, value=desc['count'])
    ws_sensor_analysis.cell(row=row_idx, column=3, value=desc['mean'])
    ws_sensor_analysis.cell(row=row_idx, column=4, value=desc['std'])
    ws_sensor_analysis.cell(row=row_idx, column=5, value=desc['min'])
    ws_sensor_analysis.cell(row=row_idx, column=6, value=desc['max'])

# Outlier summary
outlier_data = pd.DataFrame({
    'sensor_metric': ['vibration', 'temperature', 'current'],
    'outlier_count': [q1_outlier['vibration'], q1_outlier['temperature'], q1_outlier['current']],
})
outlier_row = 10
ws_sensor_analysis.cell(row=outlier_row, column=1, value='Sensor Metric')
ws_sensor_analysis.cell(row=outlier_row, column=2, value='Outlier Count')
for i, row in enumerate(outlier_data.values.tolist(), start=outlier_row + 1):
    for j, value in enumerate(row, start=1):
        ws_sensor_analysis.cell(row=i, column=j, value=value)

# Histograms / distribution charts
for idx, metric in enumerate(['vibration', 'temperature', 'current']):
    s = sensor[metric].dropna()
    hist = pd.cut(s, bins=10, duplicates='drop').value_counts().sort_index().reset_index()
    hist.columns = ['bin', 'count']
    start_col = 8 + (idx * 9)
    ws_sensor_analysis.cell(row=1, column=start_col, value=f'{metric.title()} Distribution')
    ws_sensor_analysis.cell(row=2, column=start_col, value='Bin')
    ws_sensor_analysis.cell(row=2, column=start_col + 1, value='Count')
    for i, data_row in enumerate(hist.values.tolist(), start=3):
        ws_sensor_analysis.cell(row=i, column=start_col, value=str(data_row[0]))
        ws_sensor_analysis.cell(row=i, column=start_col + 1, value=data_row[1])
    chart = BarChart()
    chart.title = f'{metric.title()} Distribution'
    chart.y_axis.title = 'Count'
    chart.x_axis.title = 'Bin'
    chart.style = 12
    chart.height = 5
    chart.width = 8
    cats = Reference(ws_sensor_analysis, min_col=start_col, min_row=3, max_row=2 + len(hist))
    data_ref = Reference(ws_sensor_analysis, min_col=start_col + 1, min_row=2, max_row=2 + len(hist))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    ws_sensor_analysis.add_chart(chart, f'{get_column_letter(start_col)}{15 + idx * 15}')

# Formatting for sensor analysis sheet
for row in ws_sensor_analysis.iter_rows(min_row=1, max_row=ws_sensor_analysis.max_row, min_col=1, max_col=ws_sensor_analysis.max_column):
    for cell in row:
        cell.border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9'),
        )
for cell in ws_sensor_analysis["A1:F1"]:
    for c in cell:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
for cell in ws_sensor_analysis["A10:B10"]:
    for c in cell:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")

# Final workbook formatting and workbook save
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = 'A2' if ws.max_row > 1 else 'A1'

wb.save(output_path)
print(f"Saved workbook to: {output_path}")

# verify workbook opens and includes all expected sheets/charts
wb_check = load_workbook(output_path, data_only=True)
print("Workbook sheets:", wb_check.sheetnames)
for ws in wb_check.worksheets:
    print(ws.title, len(ws._charts))

# Create completion report
report_lines = [
    "# Round 3 Completion Report",
    "",
    "## Deliverables",
    "- MechaCare_Round3_Final.xlsx — final workbook with management, cleaned data, machine analysis, alarm analysis, maintenance analysis, pre-maintenance analysis, load-vs-downtime, and sensor analysis sheets.",
    "- This workbook was generated from the project's cleaned CSVs and Round 2 output files and intentionally preserves raw data.",
    "",
    "## Summary of findings",
    "- Sensor duplicates were removed by business key (machine, timestamp) and missing values were retained with flags rather than filled with fabricated values.",
    "- IQR outlier counts were: vibration 720, temperature 671, current 344.",
    "- Normalized alarm frequencies were: general_warning 636, high_vibration 218, high_temperature 64, overcurrent 26, unknown 4.",
    "- The machine-level load vs downtime correlation was 0.0409, indicating a very weak relationship and no causal claim.",
    "- The strongest service pressure was in Compressor and CNC types based on average service events per machine; the dataset supports descriptive comparison rather than predictive claims.",
    "- Pre-maintenance readings were aligned to each machine's next service event using a 0–21 day window before service.",
    "",
    "## Limitations",
    "- Native Excel PivotTables are not generated by openpyxl; equivalent summary tables and charts were created instead.",
    "- Some maintenance action and downtime fields are incomplete, so the maintenance analysis remains descriptive and retrospective.",
    "- This workbook is designed to be evidence-based and submission-ready for assessment review without claiming predictive validity.",
]
(root / "ROUND3_COMPLETION_REPORT.md").write_text("\n".join(report_lines), encoding="utf-8")
print(f"Saved report to: {root / 'ROUND3_COMPLETION_REPORT.md'}")
