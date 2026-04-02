import pandas as pd

# ----------- INPUT FILES -------------
higher_file = r"C:\Users\devis\OneDrive\Desktop\rtrp\higher_education.xlsx"
final_file = r"C:\Users\devis\OneDrive\Desktop\rtrp\updated_final_sheet.xlsx"  
output_file = "final_with_higher_education.xlsx"

# ---------------- READ FINAL SHEET ----------------
final_df = pd.read_excel(final_file)

# Ensure Roll No is string
final_df["Roll No"] = final_df["Roll No"].astype(str).str.strip()

# Add Higher Education column if not present
if "Higher Education" not in final_df.columns:
    final_df["Higher Education"] = None

# ---------------- READ HIGHER EDUCATION SHEETS ----------------
xls = pd.ExcelFile(higher_file)
higher_list = []

for sheet in xls.sheet_names:
    
    # Skip 2025 sheet
    if sheet == "2025":
        continue

    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = df.columns.str.strip()

    # Convert roll number to string
    df["Roll Number"] = df["Roll Number"].astype(str).str.strip()

    # Take only required columns
    df = df[["Roll Number", "Higher Education"]]

    higher_list.append(df)

# Combine all sheets
higher_df = pd.concat(higher_list, ignore_index=True)

# Remove duplicates (no repetition of students)
higher_df = higher_df.drop_duplicates(subset="Roll Number", keep="first")

# ---------------- MERGE INTO FINAL SHEET ----------------
for _, row in higher_df.iterrows():
    roll = row["Roll Number"]
    higher = row["Higher Education"]

    mask = final_df["Roll No"] == roll

    if mask.any():
        final_df.loc[mask, "Higher Education"] = higher

# ---------------- SAVE OUTPUT ----------------
final_df.to_excel(output_file, index=False)

print("✅ Higher Education data added successfully!")