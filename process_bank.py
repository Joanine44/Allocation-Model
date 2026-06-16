import pandas as pd
import re

print("Script is running...")

# ✅ Updated file name
input_file = r"C:\JOANINE PYTHON\WEB DEV PROJECTS\Allocation Model\Input\Std Bank 4174.xlsx"
output_file = r"C:\JOANINE PYTHON\WEB DEV PROJECTS\Allocation Model\Output\processed_transactions.xlsx"

# Load Excel file
df = pd.read_excel(input_file, engine='openpyxl')

# Column containing description
ref_col = "Description"

# Pattern: 3 letters + 9 digits anywhere in text
pattern = re.compile(r'[A-Z]{3}\d{9}')

valid_rows = []
suspense_rows = []

for _, row in df.iterrows():
    description = str(row[ref_col]).upper()

    match = pattern.search(description)

    if match:
        # ✅ Extract Matter Number
        matter_number = match.group()
        row["Matter Number"] = matter_number

        # ✅ Add new column "Description 1"
        row["Description 1"] = "Std B 4174"

        # ✅ Remove original Description (only for valid)
        row_valid = row.drop(labels=[ref_col])

        valid_rows.append(row_valid)

    else:
        # ✅ Keep original Description in suspense
        row["Reason"] = "No valid Matter Number found"
        suspense_rows.append(row)

# Convert to DataFrames
valid_df = pd.DataFrame(valid_rows)
suspense_df = pd.DataFrame(suspense_rows)

# Save to Excel with 2 sheets
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    valid_df.to_excel(writer, sheet_name='Valid', index=False)
    suspense_df.to_excel(writer, sheet_name='Suspense', index=False)

print("✅ DONE")
print("Valid:", len(valid_df))
print("Suspense:", len(suspense_df))