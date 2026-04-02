import pandas as pd

file_path = r"C:\Users\devis\OneDrive\Desktop\rtrp\final_data.xlsx"
output_file = r"updated_file.xlsx"

xls = pd.ExcelFile(file_path)
updated_sheets = {}

for sheet_name in xls.sheet_names:

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    df.columns = df.columns.astype(str).str.strip()

    if "X %" in df.columns:

        # convert to string, remove % and spaces
        df["X %"] = df["X %"].astype(str).str.replace("%", "").str.strip()

        # convert to numeric
        df["X %"] = pd.to_numeric(df["X %"], errors="coerce")

        # divide by 10 if value > 10
        df.loc[df["X %"] > 10, "X %"] = df["X %"] / 10

        # multiply by 10 if value < 1
        df.loc[df["X %"] < 1, "X %"] = df["X %"] * 10


    updated_sheets[sheet_name] = df


with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for sheet_name, df in updated_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"✅ Updated Excel saved successfully as: {output_file}")
