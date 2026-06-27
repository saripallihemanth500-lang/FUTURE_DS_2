import sys
import pandas as pd
import numpy as np

INPUT_FILE = sys.argv[1] if len(sys.argv) > 1 else "WA_Fn-UseC_-Telco-Customer-Churn.csv"
OUTPUT_FILE = sys.argv[2] if len(sys.argv) > 2 else "Telco_Customer_Churn_Cleaned.xlsx"

df = pd.read_csv(INPUT_FILE)

orig_count = len(df)

for col in df.columns:
    if df[col].dtype == "object":
        df[col] = df[col].astype(str).str.strip()

removed_frames = []

dup_mask = df.duplicated(keep="first")

removed_dup = df[dup_mask].copy()
removed_dup["RemovalReason"] = "Duplicate Row"

removed_frames.append(removed_dup)

df = df[~dup_mask].copy()

cust_dup_mask = df["customerID"].duplicated(keep="first")

removed_cust_dup = df[cust_dup_mask].copy()
removed_cust_dup["RemovalReason"] = "Duplicate CustomerID"

removed_frames.append(removed_cust_dup)

df = df[~cust_dup_mask].copy()

df["TotalCharges"] = pd.to_numeric(
    df["TotalCharges"],
    errors="coerce"
)

missing_tc_mask = df["TotalCharges"].isna()

removed_missing_tc = df[missing_tc_mask].copy()
removed_missing_tc["RemovalReason"] = "Missing TotalCharges"

removed_frames.append(removed_missing_tc)

df = df[~missing_tc_mask].copy()

yes_no_cols = [
    "Partner",
    "Dependents",
    "PhoneService",
    "PaperlessBilling",
    "Churn"
]

for col in yes_no_cols:
    if col in df.columns:
        df[col] = df[col].str.title()

df["SeniorCitizenFlag"] = np.where(
    df["SeniorCitizen"] == 1,
    "Senior Citizen",
    "Non-Senior Citizen"
)

df["TenureGroup"] = pd.cut(
    df["tenure"],
    bins=[0, 12, 24, 48, 72],
    labels=[
        "0-12 Months",
        "13-24 Months",
        "25-48 Months",
        "49-72 Months"
    ],
    include_lowest=True
)

df["RevenueCategory"] = pd.cut(
    df["MonthlyCharges"],
    bins=[0, 35, 70, 120],
    labels=["Low", "Medium", "High"]
)

removed_df = pd.concat(
    removed_frames,
    ignore_index=True
)

print(f"Original Rows: {orig_count:,}")
print(f"Removed Duplicate Rows: {len(removed_dup):,}")
print(f"Removed Duplicate Customer IDs: {len(removed_cust_dup):,}")
print(f"Removed Missing TotalCharges: {len(removed_missing_tc):,}")
print(f"Final Cleaned Rows: {len(df):,}")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="xlsxwriter"
) as writer:

    summary_rows = [
        ("Original Rows", orig_count),
        ("Removed Duplicate Rows", len(removed_dup)),
        ("Removed Duplicate Customer IDs", len(removed_cust_dup)),
        ("Removed Missing TotalCharges", len(removed_missing_tc)),
        ("Final Cleaned Rows", len(df))
    ]

    pd.DataFrame(
        summary_rows,
        columns=["Metric", "Value"]
    ).to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    df.to_excel(
        writer,
        sheet_name="Cleaned_Data",
        index=False
    )

    removed_df.to_excel(
        writer,
        sheet_name="Removed_Data",
        index=False
    )

print(f"\nSaved: {OUTPUT_FILE}")
