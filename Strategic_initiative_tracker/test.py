
# ============================================
# STRATEGIC INITIATIVE TRACKER
# NYC CAPITAL PROJECTS DATASET
# ============================================

import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt

from pathlib import Path

# --------------------------------------------
# STEP 1: Define dataset path
# --------------------------------------------

file_path = Path("data/nyc_schedule_history.csv")

# Check if the file exists
if not file_path.exists():
    raise FileNotFoundError(
        "Dataset not found! Check the data folder and CSV filename."
    )

# --------------------------------------------
# STEP 2: Load the dataset
# --------------------------------------------

df = pd.read_csv(file_path)

print("\nDataset loaded successfully!")

# --------------------------------------------
# STEP 3: Display first 5 rows
# --------------------------------------------

print("\nFirst 5 records:")
print(df.head())

# --------------------------------------------
# STEP 4: Display column names
# --------------------------------------------

print("\nColumn names:")
print(df.columns.tolist())

# --------------------------------------------
# STEP 5: Display dataset shape
# --------------------------------------------

print("\nDataset shape:")
print(df.shape)

# --------------------------------------------
# STEP 6: Display data information
# --------------------------------------------

print("\nDataset information:")
df.info()

# --------------------------------------------
# STEP 7: Check missing values
# --------------------------------------------

print("\nMissing values:")
print(df.isnull().sum())

# --------------------------------------------
# STEP 8: Display statistical summary
# --------------------------------------------

print("\nStatistical summary:")
print(df.describe(include="all"))

# --------------------------------------------
# STEP 9: Display dataset sample
# --------------------------------------------

print("\nDataset inspection completed!")
print("\nDataset loaded successfully!")
print("Total rows:", len(df))
print("Total columns:", len(df.columns))

print("\nColumn names:")
print(df.columns.tolist())

print("\nFirst 5 rows:")
print(df.head())

# ==========================================
# STEP 10: DATA CLEANING
# ==========================================

print("\nStarting data cleaning...")

# 1. Create a copy of the original dataset
df = df.copy()

# 2. Standardize column names
df.columns = (
    df.columns
    .str.strip()
    .str.lower()
    .str.replace(r"[^a-z0-9]+", "_", regex=True)
    .str.strip("_")
)

# 3. Remove duplicate records
before = len(df)
df = df.drop_duplicates()
after = len(df)

print("Duplicate records removed:", before - after)

# 4. Convert date columns into proper date format
for column in ["completion_date", "data_date"]:
    if column in df.columns:
        df[column] = pd.to_datetime(
            df[column], errors="coerce"
        )

# 5. Convert numeric columns
for column in ["reporting_period", "variance_day"]:
    if column in df.columns:
        df[column] = pd.to_numeric(
            df[column], errors="coerce"
        )

# 6. Remove records without a project ID
if "pid" in df.columns:
    df = df.dropna(subset=["pid"])

# 7. Display cleaned dataset information
print("\nData cleaning completed successfully!")

print("Rows after cleaning:", len(df))
print("Columns after cleaning:", len(df.columns))

print("\nMissing values after cleaning:")
print(df.isnull().sum())

print("\nCleaned dataset preview:")
print(df.head())

# ==========================================
# STEP 11: PROJECT STATUS ANALYSIS
# ==========================================

print("\nStarting project status analysis...")

# 1. Sort projects by ID and reporting period
df = df.sort_values(
    by=["pid", "reporting_period"]
)

# 2. Select the latest reporting record for each project
latest_projects = df.drop_duplicates(
    subset=["pid"],
    keep="last"
).copy()

# 3. Calculate days remaining until completion
today = pd.Timestamp.today().normalize()

latest_projects["days_remaining"] = (
    latest_projects["completion_date"] - today
).dt.days

# 4. Create a schedule review status
latest_projects["schedule_status"] = np.select(
    [
        latest_projects["completion_date"].isna(),
        latest_projects["completion_date"] < today,
        latest_projects["completion_date"] >= today
    ],
    [
        "Date Not Available",
        "Completion Date Passed - Review",
        "Upcoming Completion"
    ],
    default="Unknown"
)

# 5. Display project status summary
print("\nProject Status Analysis Completed!")

print("\nTotal unique projects:", len(latest_projects))

print("\nProjects by schedule status:")
print(
    latest_projects["schedule_status"].value_counts()
)

print("\nProjects by managing agency:")
print(
    latest_projects["managing_agency"].value_counts().head(10)
)

print("\nLatest project records:")
print(
    latest_projects[
        [
            "pid",
            "agency_project_name",
            "managing_agency",
            "completion_date",
            "days_remaining",
            "schedule_status"
        ]
    ].head(10)
)

# ==========================================
# STEP 12: CREATE CHARTS AND EXCEL REPORT
# ==========================================

print("\nCreating charts and Excel report...")

# 1. Create output folder
output_folder = Path("outputs")
output_folder.mkdir(exist_ok=True)

# 2. Projects by managing agency
agency_counts = (
    latest_projects["managing_agency"]
    .fillna("Unknown")
    .value_counts()
    .head(10)
)

plt.figure(figsize=(12, 6))
agency_counts.plot(kind="bar")
plt.title("Top 10 Agencies by Number of Projects")
plt.xlabel("Managing Agency")
plt.ylabel("Number of Projects")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()

plt.savefig(output_folder / "projects_by_agency.png")
plt.close()

print("Agency chart created!")

# 3. Projects by current phase
phase_counts = (
    latest_projects["current_phase"]
    .fillna("Unknown")
    .value_counts()
)

plt.figure(figsize=(10, 6))
phase_counts.plot(kind="bar")
plt.title("Projects by Current Phase")
plt.xlabel("Project Phase")
plt.ylabel("Number of Projects")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()

plt.savefig(output_folder / "projects_by_phase.png")
plt.close()

print("Project phase chart created!")

# 4. Schedule variance distribution
variance_data = df["variance_day"].dropna()

plt.figure(figsize=(10, 6))
plt.hist(variance_data, bins=30)
plt.title("Distribution of Schedule Variance")
plt.xlabel("Schedule Variance (Days)")
plt.ylabel("Number of Records")
plt.tight_layout()

plt.savefig(output_folder / "schedule_variance.png")
plt.close()

print("Schedule variance chart created!")

# 5. Create summary tables
agency_summary = agency_counts.to_frame(
    name="Project Count"
)

phase_summary = phase_counts.to_frame(
    name="Project Count"
)

status_summary = (
    latest_projects["schedule_status"]
    .value_counts()
    .to_frame(name="Project Count")
)

# 6. Export data to Excel
excel_path = output_folder / "nyc_capital_projects_analysis.xlsx"

with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:

    latest_projects.to_excel(
        writer,
        sheet_name="Latest Projects",
        index=False
    )

    agency_summary.to_excel(
        writer,
        sheet_name="Agency Summary"
    )

    phase_summary.to_excel(
        writer,
        sheet_name="Phase Summary"
    )

    status_summary.to_excel(
        writer,
        sheet_name="Status Summary"
    )

    df.to_excel(
        writer,
        sheet_name="Cleaned Data",
        index=False
    )

print("\nExcel report created successfully!")
print("Report location:", excel_path)

print("\nPROJECT ANALYSIS COMPLETED!")

# ==========================================
# STEP 13: FORMAT EXCEL REPORT AND DASHBOARD
# ==========================================

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

print("\nCreating professional Excel dashboard...")

# Open the Excel report
workbook = load_workbook(excel_path)

# Create Dashboard sheet at the beginning
if "Dashboard" in workbook.sheetnames:
    del workbook["Dashboard"]

dashboard = workbook.create_sheet("Dashboard", 0)

# Dashboard title
dashboard["A1"] = "NYC CAPITAL PROJECTS DASHBOARD"
dashboard["A1"].font = Font(
    size=18, bold=True, color="FFFFFF"
)
dashboard["A1"].fill = PatternFill(
    "solid", fgColor="1F4E78"
)
dashboard.merge_cells("A1:D1")

# Dashboard metrics
dashboard["A3"] = "Total Projects"
dashboard["B3"] = len(latest_projects)

dashboard["A4"] = "Managing Agencies"
dashboard["B4"] = latest_projects["managing_agency"].nunique()

dashboard["A5"] = "Project Phases"
dashboard["B5"] = latest_projects["current_phase"].nunique()

dashboard["A6"] = "Schedule Records"
dashboard["B6"] = len(df)

# Style metric cells
for row in range(3, 7):
    dashboard[f"A{row}"].font = Font(bold=True)
    dashboard[f"A{row}"].fill = PatternFill(
        "solid", fgColor="D9EAF7"
    )

    dashboard[f"B{row}"].font = Font(
        bold=True, color="1F4E78"
    )

# Add project status summary
dashboard["A8"] = "Schedule Status"
dashboard["B8"] = "Project Count"

dashboard["A8"].font = Font(bold=True, color="FFFFFF")
dashboard["B8"].font = Font(bold=True, color="FFFFFF")

dashboard["A8"].fill = PatternFill(
    "solid", fgColor="1F4E78"
)
dashboard["B8"].fill = PatternFill(
    "solid", fgColor="1F4E78"
)

for row_number, (status, count) in enumerate(
    latest_projects["schedule_status"].value_counts().items(),
    start=9
):
    dashboard.cell(row=row_number, column=1, value=status)
    dashboard.cell(row=row_number, column=2, value=count)

# Create status bar chart
chart = BarChart()
chart.title = "Projects by Schedule Status"
chart.y_axis.title = "Number of Projects"
chart.x_axis.title = "Schedule Status"

data = Reference(
    dashboard,
    min_col=2,
    min_row=8,
    max_row=8 + latest_projects["schedule_status"].nunique()
)

categories = Reference(
    dashboard,
    min_col=1,
    min_row=9,
    max_row=8 + latest_projects["schedule_status"].nunique()
)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)
chart.height = 8
chart.width = 16

dashboard.add_chart(chart, "D3")

# Format all report sheets
for sheet in workbook.worksheets:

    sheet.freeze_panes = "A2"

    if sheet.max_row > 1:
        sheet.auto_filter.ref = sheet.dimensions

    for cell in sheet[1]:
        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )
        cell.alignment = Alignment(
            horizontal="center"
        )

    # Automatically adjust column widths
    for column_cells in sheet.columns:
        max_length = 0
        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        sheet.column_dimensions[column_letter].width = min(
            max_length + 2, 35
        )

# Save formatted workbook
workbook.save(excel_path)

print("\nProfessional Excel dashboard created!")
print("Saved at:", excel_path)
