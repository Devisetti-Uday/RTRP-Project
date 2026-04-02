import pandas as pd

# -------- FILE PATH --------
file1 = r"C:\Users\devis\OneDrive\Desktop\rtrp\dataset\Results\2020-24\sem8.xlsx"
output = "batch_sem.xlsx"

# -------- INPUTS --------
batch = "2020-2024"
semester = 8

# -------- SUBJECT CODE → SUBJECT NAME MAP --------
subjects = [					
    ("GR20A4051","Natural Language Processing (Minors)"),					
    ("GR20A4084","Pavement Design"),					
    ("GR20A4089","Green Building Technology"),					
    ("GR20A4091","Entrepreneurship and Project Management"),					
    ("GR20A4130","Project Work - Phase II"),					
    ("GR20A6001","Plastic Waste Management (through MOOCs)"),					
    ("GR20A6008","Safety in Construction (through MOOCs)"),					
    ("GR20A6017","Introduction to Civil Engineering Profession (through MOOCs)"),					
    ("GR20A4051","Natural Language Processing (Minors)"),						
    ("GR20A4092","Power System Monitoring and Control"),						
    ("GR20A4096","Industrial IoT"),						
    ("GR20A4098","Electric Smart Grid"),						
    ("GR20A4130","Project Work - Phase II"),						
    ("GR20A6004","Introduction to Internet of Things (through MOOCs)"),						
    ("GR20A6007","Cloud Computing (through MOOCs)"),						
    ("GR20A3067","Augmented Reality and Virtual Reality"),							
    ("GR20A4051","Natural Language Processing (Minors)"),							
    ("GR20A4100","Rapid Prototyping and Tooling"),							
    ("GR20A4104","Production Planning and Control"),							
    ("GR20A4130","Project Work - Phase II"),							
    ("GR20A6006","Robotics (through MOOCs)"),							
    ("GR20A6011","Corrosion Protection method (through MOOCs)"),							
    ("GR20A6018","Introduction to Industry 4.0 and Industrial Internet of Things(through MOOCs)"), 							
    ("GR20A3118","Cloud Computing"),						
    ("GR20A4051","Natural Language Processing (Minors)"),						
    ("GR20A4107","Satellite Communication"),						
    ("GR20A4111","Global Navigation Satellite System"),						
    ("GR20A4130","Project Work - Phase II"),						
    ("GR20A6002","Data Science for Engineers (through MOOCs)"),						
    ("GR20A6005","The Joy of Computing using Python (through MOOCs)"),						
    ("GR20A3140","Fundamentals of Management and Entrepreneurship"),						
    ("GR20A4067","Human Computer Interaction"),						
    ("GR20A4115","Cyber Security"),						
    ("GR20A4130","Project Work - Phase II"),						
    ("GR20A4144","Technical Paper writing (Honors)"),						
    ("GR20A6002","Data Science for Engineers (through MOOCs)"),						
    ("GR20A6009","Data Analytics with Python (through MOOCs)"),						
    ("GR20A4119","Software Project Management"),							
    ("GR20A4120","E-Commerce"),							
    ("GR20A4124","Design Patterns"),							
    ("GR20A4130","Project Work - Phase II"),							
    ("GR20A6003","Big Data Computing (through MOOCs)"),							
    ("GR20A6010","User-Centric Computing for Human Interaction (through MOOCs)"),							
    ("GR20A4130","Project Work - Phase II"),			
    ("GR20A4133","Psychology"),			
    ("GR20A4134","Enterprise Systems"),			
    ("GR20A4138","IT Project Management"),			
    ("GR20A3140","Fundamentals of Management and Entrepreneurship"),					
    ("GR20A4118","Software Product Development and Management"),					
    ("GR20A4125","Cyber Forensics"),					
    ("GR20A4130","Project Work - Phase II"),					
    ("GR20A6004","Introduction to Internet of Things (through MOOCs)"),					
    ("GR20A6010","User-Centric Computing for Human Interaction (through MOOCs)"),					
    ("GR20A3140","Fundamentals of Management and Entrepreneurship"),					
    ("GR20A4058","Software Testing Methodologies"),					
    ("GR20A4115","Cyber Security"),					
    ("GR20A4130","Project Work - Phase II"),					
    ("GR20A6010","User-Centric Computing for Human Interaction (through MOOCs)")					
]


subject_map = dict(subjects)

# -------- READ FILE --------
xls = pd.ExcelFile(file1)

all_data = []

for sheet in xls.sheet_names:

    # skip unwanted sheets
    if "dont" in sheet.lower():
        continue

    # -------- READ FROM ROW 10 --------
    df = pd.read_excel(
        xls,
        sheet_name=sheet,
        header=9   # row 10 as header (0-indexed)
    )

    # remove column A
    df = df.iloc[:, 1:]

    # clean column names
    df.columns = df.columns.astype(str).str.strip()

    # skip sheets without roll number
    if "Roll No" not in df.columns:
        continue

    # remove SGPA and Credits
    df = df.drop(
        columns=[col for col in df.columns if col.lower() in ["sgpa", "credits"]],
        errors="ignore"
    )

    # -------- WIDE → LONG --------
    melted = df.melt(
        id_vars=["Roll No"],
        var_name="SubjectCode",
        value_name="Grade"
    )

    # -------- MAP CODE → SUBJECT NAME --------
    melted["Subject"] = melted["SubjectCode"].map(subject_map)

    # add metadata
    melted["Branch"] = sheet
    melted["Batch"] = batch
    melted["Semester"] = semester

    # rename column
    melted = melted.rename(columns={"Roll No": "Student"})

    # final columns
    melted = melted[
        ["Branch", "Batch", "Semester", "Student","SubjectCode", "Subject", "Grade"]
    ]

    all_data.append(melted)

# combine all sheets
final_df = pd.concat(all_data, ignore_index=True)

# save
final_df.to_excel(output, index=False)

print("✅ batch_sem.xlsx created successfully!")