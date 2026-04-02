import pandas as pd

sheet1_path = r"C:\Users\devis\OneDrive\Desktop\rtrp\final_data_2020.xlsx"
sheet2_path = r"C:\Users\devis\OneDrive\Desktop\rtrp\Results-batchwise\2016-2020.xlsx"
output_path = "final_output.xlsx"

# -------- LOAD SHEET 1 --------
df1 = pd.read_excel(sheet1_path)
df1.columns = df1.columns.astype(str).str.strip()

# Insert CGPA & Division at column L & M
if "CGPA" not in df1.columns:
    df1.insert(11, "CGPA", "")
if "Division" not in df1.columns:
    df1.insert(12, "Division", "")

xls2 = pd.ExcelFile(sheet2_path)

# -------- PROCESS EACH ROW --------
for i in range(len(df1)):

    roll1 = str(df1.loc[i, "Roll No"]).strip()
    branch = str(df1.loc[i, "Branch"]).strip()

    if branch not in xls2.sheet_names:
        continue

    # Read raw to find header row
    df_raw = pd.read_excel(sheet2_path, sheet_name=branch, header=None)

    header_row = None
    for r in range(len(df_raw)):
        if df_raw.iloc[r].astype(str).str.contains("Roll", case=False).any():
            header_row = r
            break

    if header_row is None:
        continue

    df2 = pd.read_excel(sheet2_path, sheet_name=branch, header=header_row)
    df2.columns = df2.columns.astype(str).str.strip()

    # Convert column names to uppercase for safe matching
    cols_upper = [c.upper() for c in df2.columns]

    roll_col = None
    cgpa_col = None
    div_col = None

    for idx, col in enumerate(cols_upper):
        if "ROLL" in col:
            roll_col = df2.columns[idx]
        if "CGPA" in col:
            cgpa_col = df2.columns[idx]
        if "DIV" in col:
            div_col = df2.columns[idx]

    if not roll_col or not cgpa_col or not div_col:
        continue

    # Clean roll numbers
    df2[roll_col] = (
        df2[roll_col]
        .astype(str)
        .str.replace("\xa0", "", regex=False)
        .str.strip()
    )

    match = df2[df2[roll_col] == roll1]

    if not match.empty:
        df1.loc[i, "CGPA"] = match.iloc[0][cgpa_col]
        df1.loc[i, "Division"] = match.iloc[0][div_col]

# -------- SAVE --------
df1.to_excel(output_path, index=False)

print("✅ CGPA & Division added successfully!")