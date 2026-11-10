import pandas as pd
import numpy as np

# --------------------------------
# 1. LOAD THE DATASET
# --------------------------------

file_path = "Warangal_Weather_1000_Days.xlsx"

df = pd.read_excel(file_path)

print("ORIGINAL DATASET")
print(df.head())

print("\nDataset Shape:", df.shape)


# --------------------------------
# 2. STANDARDIZE COLUMN NAMES
# --------------------------------

df.columns = df.columns.str.strip().str.lower()

# Workbooks may encode degree symbols differently, so map columns by their
# stable prefix rather than unit suffixes.
rename_map = {}
prefixes = {
    "max temperature": "maximum_temperature",
    "min temperature": "minimum_temperature",
    "average temperature": "average_temperature",
    "average humidity": "average_humidity",
    "rainfall": "rainfall",
    "max wind speed": "maximum_wind_speed",
    "date": "date",
}
for col in df.columns:
    for prefix, canonical_name in prefixes.items():
        if col.startswith(prefix):
            rename_map[col] = canonical_name
            break

df = df.rename(columns=rename_map)
df = df.drop(
    columns=[col for col in df.columns if str(col).startswith("unnamed:")],
)

print("\nColumn Names:")
print(df.columns.tolist())


# --------------------------------
# 3. HANDLE DATE FORMATS
# --------------------------------

df["date"] = pd.to_datetime(
    df["date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

# Remove invalid dates
df = df.dropna(subset=["date"])


# --------------------------------
# 4. REMOVE DUPLICATE RECORDS
# --------------------------------

print("\nDuplicate Records:", df.duplicated().sum())

df = df.drop_duplicates()

# Remove duplicate dates
df = df.drop_duplicates(subset=["date"], keep="first")


# --------------------------------
# 5. CONVERT NUMERIC COLUMNS
# --------------------------------

numeric_columns = [
    "maximum_temperature",
    "minimum_temperature",
    "average_temperature",
    "average_humidity",
    "rainfall",
    "maximum_wind_speed"
]

for col in numeric_columns:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")


# --------------------------------
# 6. HANDLE MISSING VALUES
# --------------------------------

print("\nMissing Values Before Cleaning:")
print(df.isnull().sum())

for col in numeric_columns:
    if col in df.columns:
        df[col] = df[col].fillna(df[col].median())


# --------------------------------
# 7. CHECK INVALID VALUES
# --------------------------------

# Humidity must be between 0 and 100
if "average_humidity" in df.columns:
    df.loc[
        ~df["average_humidity"].between(0, 100),
        "average_humidity"
    ] = np.nan

# Rainfall cannot be negative
if "rainfall" in df.columns:
    df.loc[df["rainfall"] < 0, "rainfall"] = np.nan

# Wind speed cannot be negative
if "maximum_wind_speed" in df.columns:
    df.loc[df["maximum_wind_speed"] < 0, "maximum_wind_speed"] = np.nan

# Temperature consistency
if "maximum_temperature" in df.columns and "minimum_temperature" in df.columns:
    invalid_temp = (
        df["maximum_temperature"] < df["minimum_temperature"]
    )

    df.loc[invalid_temp, [
        "maximum_temperature",
        "minimum_temperature"
    ]] = np.nan

# Fill newly created missing values
for col in numeric_columns:
    if col in df.columns:
        df[col] = df[col].fillna(df[col].median())


# --------------------------------
# 8. CREATE MONTH COLUMN
# --------------------------------

df["month"] = df["date"].dt.strftime("%B")


# --------------------------------
# 9. SORT DATA BY DATE
# --------------------------------

df = df.sort_values(by="date")

df = df.reset_index(drop=True)


# --------------------------------
# 10. FINAL DATA VALIDATION
# --------------------------------

print("\nCLEANED DATASET")
print(df.head())

print("\nMissing Values After Cleaning:")
print(df.isnull().sum())

print("\nDuplicate Records After Cleaning:")
print(df.duplicated().sum())

print("\nFinal Dataset Shape:")
print(df.shape)


# --------------------------------
# 11. EXPORT CLEANED DATA
# --------------------------------

df.to_excel(
    "Warangal_Weather_Cleaned.xlsx",
    index=False
)

df.to_csv(
    "Warangal_Weather_Cleaned.csv",
    index=False
)

print("\nData Cleaning Completed Successfully!")

print("Cleaned Excel and CSV files saved.")