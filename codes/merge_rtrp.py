import pandas as pd

file_path = r"C:\Users\devis\OneDrive\Desktop\rtrp\Stu data\2017-21.xlsx"

required_columns = [
    "Roll No",
    "Lateral",
    "Gender",
    "Caste",
    "com District",
    "X %",
    "Inter %",
    "Admission",
    "Category",
    "Fee Reimburse"
]

xls = pd.ExcelFile(file_path)

combined_data = []

for sheet_name in xls.sheet_names:

    # read sheet with header always at row 1
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

    df.columns = df.columns.astype(str).str.strip()

    # add branch column
    df["Branch"] = sheet_name

    cols_to_keep = [col for col in required_columns if col in df.columns]
    cols_to_keep.append("Branch")

    df = df[cols_to_keep]

    combined_data.append(df)

final_df = pd.concat(combined_data, ignore_index=True)

final_df.to_excel("final_data_2021.xlsx", index=False)

print("✅ All sheets combined successfully with required columns!")
