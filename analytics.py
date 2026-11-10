from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "Warangal_Weather_Cleaned.xlsx"
REPORT_FILE = BASE_DIR / "Warangal_Weather_Analysis.xlsx"

NUMERIC_COLUMNS = [
    "maximum_temperature",
    "minimum_temperature",
    "average_temperature",
    "average_humidity",
    "rainfall",
    "maximum_wind_speed",
]


def add_dataframe(writer, dataframe, sheet_name, index=False):
    dataframe.to_excel(writer, sheet_name=sheet_name, index=index)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "B2" if index else "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column in worksheet.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 30)
        worksheet.column_dimensions[column[0].column_letter].width = width


df = pd.read_excel(DATA_FILE)
df["date"] = pd.to_datetime(df["date"])
df["year"] = df["date"].dt.year
df["month"] = df["date"].dt.month_name()

numeric_summary = df[NUMERIC_COLUMNS].describe().T.reset_index().rename(columns={"index": "metric"})

missing_values = (
    df.isna()
    .sum()
    .rename("missing_count")
    .reset_index()
    .rename(columns={"index": "column"})
)
missing_values["missing_percent"] = (
    missing_values["missing_count"] / len(df) * 100
).round(2)

monthly_summary = (
    df.groupby("month", sort=False)[NUMERIC_COLUMNS]
    .agg(["mean", "min", "max", "sum"])
    .round(2)
    .reset_index()
)
monthly_summary.columns = ["_".join(filter(None, col)).strip("_") for col in monthly_summary.columns]
month_order = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
monthly_summary["month_order"] = monthly_summary["month"].map(
    {name: index for index, name in enumerate(month_order)}
)
monthly_summary = monthly_summary.sort_values("month_order").drop(columns="month_order")

yearly_summary = (
    df.groupby("year")[NUMERIC_COLUMNS]
    .agg(["mean", "min", "max", "sum"])
    .round(2)
    .reset_index()
)
yearly_summary.columns = ["_".join(filter(None, col)).strip("_") for col in yearly_summary.columns]

correlation = df[NUMERIC_COLUMNS].corr().round(3).reset_index().rename(columns={"index": "metric"})

weather_summary = (
    df.groupby("weather code", dropna=False)
    .agg(days=("date", "count"))
    .reset_index()
    .sort_values("days", ascending=False)
)
weather_summary["percent"] = (weather_summary["days"] / len(df) * 100).round(2)

extreme_rows = []
for column in NUMERIC_COLUMNS:
    series = df[column].dropna()
    extreme_rows.append({
        "metric": column,
        "highest_value": series.max(),
        "highest_date": df.loc[df[column] == series.max(), "date"].min(),
        "lowest_value": series.min(),
        "lowest_date": df.loc[df[column] == series.min(), "date"].min(),
    })
extremes = pd.DataFrame(extreme_rows)

overview = pd.DataFrame([
    ["Dataset", DATA_FILE.name],
    ["Rows", len(df)],
    ["Columns", len(df.columns)],
    ["Date range", f"{df['date'].min():%Y-%m-%d} to {df['date'].max():%Y-%m-%d}"],
    ["Weather days recorded", df["date"].nunique()],
    ["Total rainfall", round(df["rainfall"].sum(), 2)],
    ["Average temperature", round(df["average_temperature"].mean(), 2)],
    ["Average humidity", round(df["average_humidity"].mean(), 2)],
])
overview.columns = ["analysis_item", "value"]

with pd.ExcelWriter(REPORT_FILE, engine="openpyxl") as writer:
    add_dataframe(writer, overview, "Overview")
    add_dataframe(writer, missing_values, "Missing Values")
    add_dataframe(writer, numeric_summary, "Numeric Summary")
    add_dataframe(writer, monthly_summary, "Monthly Analysis")
    add_dataframe(writer, yearly_summary, "Yearly Analysis")
    add_dataframe(writer, correlation, "Correlation")
    add_dataframe(writer, weather_summary, "Weather Codes")
    add_dataframe(writer, extremes, "Weather Extremes")
    add_dataframe(writer, df, "Cleaned Data", index=False)

print(f"Analysis completed successfully: {REPORT_FILE}")