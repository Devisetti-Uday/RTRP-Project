import pandas as pd

# ----------- INPUT FILES -------------
placement_file = r"C:\Users\devis\OneDrive\Desktop\rtrp\placement.xlsx"
final_file = r"C:\Users\devis\OneDrive\Desktop\rtrp\final_data.xlsx"
output_file = "updated_final_sheet.xlsx"

# ---------------- READ FINAL SHEET ----------------
final_df = pd.read_excel(final_file)

# Ensure Roll No is string (IMPORTANT FIX)
final_df["Roll No"] = final_df["Roll No"].astype(str).str.strip()

# Add columns if not present
if "Company" not in final_df.columns:
    final_df["Company"] = None

if "Salary" not in final_df.columns:
    final_df["Salary"] = None

# ---------------- READ PLACEMENT SHEETS ----------------
xls = pd.ExcelFile(placement_file)
placement_list = []

for sheet in xls.sheet_names:
    
    # Skip 2025 sheet
    if sheet == "2025":
        continue
    
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = df.columns.str.strip()

    # Ensure Roll Number is string (IMPORTANT FIX)
    df["Roll Number"] = df["Roll Number"].astype(str).str.strip()

    # Select required columns
    df = df[["Roll Number", "Company", "Salary"]]

    placement_list.append(df)

# Combine all sheets
placement_df = pd.concat(placement_list, ignore_index=True)

# ---------------- KEEP ONLY HIGHEST SALARY PER STUDENT ----------------
placement_df["Salary"] = pd.to_numeric(placement_df["Salary"], errors="coerce")

placement_df = placement_df.sort_values("Salary", ascending=False)
placement_df = placement_df.drop_duplicates(subset="Roll Number", keep="first")

# ---------------- MERGE INTO FINAL SHEET ----------------
for _, row in placement_df.iterrows():
    roll = row["Roll Number"]
    company = row["Company"]
    salary = row["Salary"]

    mask = final_df["Roll No"] == roll

    if mask.any():
        existing_salary = final_df.loc[mask, "Salary"].values[0]

        if pd.isna(existing_salary) or salary > existing_salary:
            final_df.loc[mask, "Company"] = company
            final_df.loc[mask, "Salary"] = salary

# ---------------- SAVE OUTPUT ----------------
final_df.to_excel(output_file, index=False)

print("✅ Updated file saved successfully!")