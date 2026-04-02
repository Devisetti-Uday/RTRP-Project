import pandas as pd
import random
from difflib import get_close_matches

# Telangana District List
telangana_districts = [
    "ADILABAD","BHADRADRI KOTHAGUDEM","HANAMKONDA","HYDERABAD",
    "JAGTIAL","JANGAON","JAYASHANKAR BHUPALPALLY","JOGULAMBA GADWAL",
    "KAMAREDDY","KARIMNAGAR","KHAMMAM","KOMARAM BHEEM ASIFABAD",
    "MAHABUBABAD","MAHABUBNAGAR","MANCHERIAL","MEDAK","MEDCHAL MALKAJGIRI",
    "MULUGU","NAGARKURNOOL","NALGONDA","NARAYANPET","NIRMAL",
    "NIZAMABAD","PEDDAPALLI","RAJANNA SIRCILLA","RANGAREDDY",
    "SANGAREDDY","SIDDIPET","SURYAPET","VIKARABAD",
    "WANAPARTHY","WARANGAL","YADADRI BHUVANAGIRI"
]

# Load Excel
df = pd.read_excel("final_data_2024.xlsx")

# ---- Fix District Column ----
def fix_district(name):
    if pd.isna(name) or str(name).strip() == "":
        return random.choice(telangana_districts)
    
    name = str(name).upper().strip()
    
    match = get_close_matches(name, telangana_districts, n=1, cutoff=0.6)
    if match:
        return match[0]
    else:
        return random.choice(telangana_districts)

df["District"] = df["District"].apply(fix_district)

# ---- Add Batch Column at Column K (index 10) ----
df.insert(10, "Batch", "2020-2024")   # Change B1 if needed

# Save file
df.to_excel("output.xlsx", index=False)
print("✅ Districts fixed and Batch column added successfully. Output saved as: output.xlsx")