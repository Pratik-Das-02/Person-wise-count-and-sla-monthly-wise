# ============================================================

# FINAL SLA PERSON + INSTANCE REPORT

# GOOGLE COLAB

# ============================================================



# ============================================================

# FINAL LOGIC

#

# 1. Exact Assigned To:

#

#       Abhinav

#           ->

#       Abhinav (W)

#

#

# 2. Assigned To with shift:

#

#       Abhinav Morning

#       Abhinav Afternoon

#       Abhinav Evening

#       Abhinav Night

#

#       Abhinav Shift A

#       Abhinav Shift B

#       Abhinav Shift C

#       Abhinav Shift D

#       Abhinav Shift E

#       Abhinav Shift F

#       Abhinav Shift G

#

#           ->

#       Abhinav

#

#

# 3. All shifts are combined.

#

# 4. Exact "Abhinav" remains separate as "Abhinav (W)".

#

# 5. Count = INSTANCE-WISE

#

# 6. Average Actual SLA = INSTANCE-WISE

#

# 7. Person header has Count only.

#    Person header Average SLA is BLANK.

#

# 8. Actual SLA blank:

#       Actual SLA = Overall SLA - 1 minute

#

# 9. Average SLA output is directly HH:MM:SS text.

#

# ============================================================





# ============================================================

# 1. INSTALL

# ============================================================



!pip -q install pandas openpyxl xlrd





# ============================================================

# 2. IMPORT

# ============================================================



import pandas as pd

import numpy as np

import re



from datetime import datetime, time, timedelta



from google.colab import files



from openpyxl import load_workbook

from openpyxl.styles import Font, PatternFill, Alignment

from openpyxl.utils import get_column_letter





# ============================================================

# 3. UPLOAD EXCEL

# ============================================================



print("=" * 75)

print("UPLOAD YOUR SLA EXCEL FILE")

print("=" * 75)



uploaded = files.upload()



if not uploaded:

    raise Exception("No Excel file uploaded.")



input_file = list(uploaded.keys())[0]



print("\nUploaded file:")

print(input_file)





# ============================================================

# 4. FIND SUBMISSIONS HISTORY SHEET

# ============================================================



excel_file = pd.ExcelFile(input_file)



sheet_names = excel_file.sheet_names



print("\nAvailable sheets:")



for sheet in sheet_names:

    print(" -", sheet)





def clean_text(value):



    if pd.isna(value):

        return ""



    value = str(value).strip()



# Remove multiple spaces

    value = re.sub(

        r'\s+',

        ' ',

        value

    )



    return value





def normalize_key(value):



    value = clean_text(value)



    return re.sub(

        r'[^a-z0-9]',

        '',

        value.lower()

    )





normalized_sheets = {

    normalize_key(sheet): sheet

    for sheet in sheet_names

}





selected_sheet = None





for possible_sheet in [

    "submissionhistory",

    "submissionshistory"

]:



    if possible_sheet in normalized_sheets:



        selected_sheet = normalized_sheets[

            possible_sheet

        ]



        break





if selected_sheet is None:



    raise Exception(

        "\nERROR:\n"

        "Could not find 'Submission History' "

        "or 'Submissions History'.\n"

        "The code will NOT use another sheet."

    )





print(

    "\nSelected sheet:",

    selected_sheet

)





# ============================================================

# 5. READ SHEET

# ============================================================



df = pd.read_excel(

    input_file,

    sheet_name=selected_sheet

)



original_rows = len(df)



print(

    f"\nOriginal rows: {original_rows:,}"

)





# ============================================================

# 6. SHOW COLUMNS

# ============================================================



print("\n" + "=" * 75)

print("COLUMNS FOUND IN EXCEL")

print("=" * 75)



for i, col in enumerate(

    df.columns,

    start=1

):



    print(

        f"{i}. [{col}]"

    )





# ============================================================

# 7. FIND COLUMN

# ============================================================



column_lookup = {

    normalize_key(col): col

    for col in df.columns

}





def find_column(candidates):



# --------------------------------------------------------

# Exact match

# --------------------------------------------------------



    for candidate in candidates:



        key = normalize_key(candidate)



        if key in column_lookup:



            return column_lookup[key]





# --------------------------------------------------------

# Partial match

# --------------------------------------------------------



    for actual_key, actual_column in column_lookup.items():



        for candidate in candidates:



            candidate_key = normalize_key(

                candidate

            )



            if candidate_key in actual_key:



                return actual_column





    return None





# ============================================================

# 8. DETECT REQUIRED COLUMNS

# ============================================================



assign_col = find_column([

    "Assigned To",

    "Assign To",

    "AssignTo",

    "Assignee"

])





instance_col = find_column([

    "Instance",

    "Instance Name"

])





actual_sla_col = find_column([

    "Actual SLA",

    "ActualSLA",

    "Actual SLA Time"

])





overall_sla_col = find_column([

    "Overall SLA",

    "OverallSLA",

    "Overall SLA Time",

    "SLA"

])





print("\n" + "=" * 75)

print("DETECTED COLUMNS")

print("=" * 75)



print(

    "Assigned To :",

    assign_col

)



print(

    "Instance    :",

    instance_col

)



print(

    "Actual SLA  :",

    actual_sla_col

)



print(

    "Overall SLA :",

    overall_sla_col

)





# ============================================================

# 9. CHECK REQUIRED COLUMNS

# ============================================================



missing_columns = []





if assign_col is None:



    missing_columns.append(

        "Assigned To"

    )





if instance_col is None:



    missing_columns.append(

        "Instance"

    )





if actual_sla_col is None:



    missing_columns.append(

        "Actual SLA"

    )





if overall_sla_col is None:



    missing_columns.append(

        "Overall SLA / SLA"

    )





if missing_columns:



    raise Exception(

        "\nERROR: Required columns were not found:\n\n"

        +

        "\n".join(

            " - " + column

            for column in missing_columns

        )

    )





# ============================================================

# 10. CLEAN ASSIGNED TO / INSTANCE

# ============================================================



df["_AssignedTo"] = df[

    assign_col

].apply(clean_text)





df["_Instance"] = df[

    instance_col

].apply(clean_text)





# ============================================================

# 11. REMOVE BLANK / ZERO INSTANCE

# ============================================================



before_instance_filter = len(df)





df = df[

    (df["_Instance"] != "") &

    (

        df["_Instance"]

        .str.lower()

        != "0"

    )

].copy()





removed_blank_instance = (

    before_instance_filter

    - len(df)

)





# ============================================================

# 12. CLASSIFY ASSIGNED TO

# ============================================================



# ============================================================

# VALID SHIFT TYPES

#

# Normal shifts:

# Morning

# Afternoon

# Evening

# Night

#

# Shift-based shifts:

# Shift A

# Shift B

# Shift C

# Shift D

# Shift E

# Shift F

# Shift G

# ============================================================



VALID_SIMPLE_SHIFTS = {

    "morning",

    "afternoon",

    "evening",

    "night"

}





VALID_SHIFT_LETTERS = {

    "a",

    "b",

    "c",

    "d",

    "e",

    "f",

    "g"

}





def classify_assigned_to(assign_to):



    assign_to = clean_text(

        assign_to

    )





# --------------------------------------------------------

# Blank Assigned To

# --------------------------------------------------------



    if assign_to == "":

        return None, None





    assign_lower = assign_to.lower()





# --------------------------------------------------------

# Invalid values

# --------------------------------------------------------



    if assign_lower in [

        "0",

        "nan",

        "none"

    ]:

        return None, None





    parts = assign_to.split()





# ========================================================

# CASE 1

# NAME + SIMPLE SHIFT

#

# Examples:

#

# Abhinav Morning

# Abhinav Afternoon

# Abhinav Evening

# Abhinav Night

#

# Result:

#

# Abhinav

# ========================================================



    if len(parts) >= 2:



        last_word = parts[-1].lower()





        if last_word in VALID_SIMPLE_SHIFTS:



            person_name = " ".join(

                parts[:-1]

            ).strip()





            if person_name:



                return (

                    person_name,

                    "SHIFT"

                )





# ========================================================

# CASE 2

# NAME + SHIFT + LETTER

#

# Examples:

#

# Abhinav Shift A

# Abhinav Shift B

# Abhinav Shift C

# Abhinav Shift D

# Abhinav Shift E

# Abhinav Shift F

# Abhinav Shift G

#

# Result:

#

# Abhinav

# ========================================================



    if len(parts) >= 3:



        second_last_word = parts[-2].lower()

        last_word = parts[-1].lower()





        if (

            second_last_word == "shift"

            and

            last_word in VALID_SHIFT_LETTERS

        ):



            person_name = " ".join(

                parts[:-2]

            ).strip()





            if person_name:



                return (

                    person_name,

                    "SHIFT"

                )





# ========================================================

# EXACT NAME

#

# Example:

#

# Abhinav

#

# Result:

#

# Abhinav (W)

# ========================================================



    return (

        assign_to,

        "WORKER"

    )





classified = df[

    "_AssignedTo"

].apply(

    classify_assigned_to

)





df["_BasePerson"] = classified.apply(

    lambda x: x[0]

)





df["_AssignmentType"] = classified.apply(

    lambda x: x[1]

)





# ============================================================

# 13. REMOVE INVALID ASSIGN TO

# ============================================================



before_assignment_filter = len(df)





df = df[

    df["_BasePerson"].notna() &

    df["_AssignmentType"].notna()

].copy()





removed_invalid_assign = (

    before_assignment_filter

    - len(df)

)





# ============================================================

# 14. CREATE FINAL OUTPUT PERSON

# ============================================================



# ------------------------------------------------------------

# Exact:

#

# Abhinav

#       ->

# Abhinav (W)

#

#

# Shift:

#

# Abhinav Morning

# Abhinav Afternoon

# Abhinav Evening

# Abhinav Night

#

# Abhinav Shift A

# Abhinav Shift B

# Abhinav Shift C

# Abhinav Shift D

# Abhinav Shift E

# Abhinav Shift F

# Abhinav Shift G

#

#       ->

# Abhinav

# ------------------------------------------------------------





df["_OutputPerson"] = np.where(



    df["_AssignmentType"] == "WORKER",



    df["_BasePerson"].str.strip()

    + " (W)",



    df["_BasePerson"].str.strip()

)





# ============================================================

# 15. SLA PARSER

# ============================================================



def parse_sla_to_seconds(value):



    if pd.isna(value):



        return np.nan





# --------------------------------------------------------

# Pandas Timedelta

# --------------------------------------------------------



    if isinstance(

        value,

        pd.Timedelta

    ):



        return value.total_seconds()





# --------------------------------------------------------

# Python timedelta

# --------------------------------------------------------



    if isinstance(

        value,

        timedelta

    ):



        return value.total_seconds()





# --------------------------------------------------------

# datetime.time

# --------------------------------------------------------



    if isinstance(

        value,

        time

    ):



        return (

            value.hour * 3600

            +

            value.minute * 60

            +

            value.second

            +

            value.microsecond / 1000000

        )





# --------------------------------------------------------

# datetime

# --------------------------------------------------------



    if isinstance(

        value,

        datetime

    ):



        return (

            value.hour * 3600

            +

            value.minute * 60

            +

            value.second

            +

            value.microsecond / 1000000

        )





# --------------------------------------------------------

# Numeric Excel time

# --------------------------------------------------------



    if isinstance(

        value,

        (

            int,

            float,

            np.integer,

            np.floating

        )

    ):



        if pd.isna(value):



            return np.nan





        number = float(value)





# Excel time stored as fraction of day

        if 0 <= number < 10:



            return (

                number

                * 86400

            )





        return np.nan





# --------------------------------------------------------

# STRING

# --------------------------------------------------------



    text = str(value).strip()





    if text == "":

        return np.nan





# Remove "days"

    text = re.sub(

        r'^\d+\s+days?\s+',

        '',

        text,

        flags=re.IGNORECASE

    )





# --------------------------------------------------------

# HH:MM:SS

# --------------------------------------------------------



    match = re.match(

        r'^(\d+):(\d{1,2}):(\d{1,2})(?:\\.(\d+))?$',

        text

    )





    if match:



        hours = int(

            match.group(1)

        )



        minutes = int(

            match.group(2)

        )



        seconds = int(

            match.group(3)

        )



        fraction = match.group(4)



        microseconds = 0





        if fraction:



            microseconds = int(

                (fraction + "000000")[:6]

            )





        return (

            hours * 3600

            +

            minutes * 60

            +

            seconds

            +

            microseconds / 1000000

        )





# --------------------------------------------------------

# MM:SS

# --------------------------------------------------------



    match = re.match(

        r'^(\d+):(\d{1,2})$',

        text

    )





    if match:



        minutes = int(

            match.group(1)

        )



        seconds = int(

            match.group(2)

        )





        return (

            minutes * 60

            +

            seconds

        )





# --------------------------------------------------------

# Pandas timedelta

# --------------------------------------------------------



    try:



        td = pd.to_timedelta(

            text

        )





        if not pd.isna(td):



            return td.total_seconds()





    except:



        pass





    return np.nan





# ============================================================

# 16. PARSE OVERALL SLA

# ============================================================



df["_OverallSLA_Seconds"] = df[

    overall_sla_col

].apply(

    parse_sla_to_seconds

)





# ============================================================

# 17. PARSE ACTUAL SLA

# ============================================================



df["_ActualSLA_Seconds"] = df[

    actual_sla_col

].apply(

    parse_sla_to_seconds

)





# ============================================================

# 18. FALLBACK

#

# Actual SLA blank:

#

# Overall SLA - 1 minute

# ============================================================



actual_missing = (

    df["_ActualSLA_Seconds"].isna()

)





fallback_rows = (



    actual_missing



    &



    df["_OverallSLA_Seconds"].notna()

)





df.loc[

    fallback_rows,

    "_ActualSLA_Seconds"

] = (



    df.loc[

        fallback_rows,

        "_OverallSLA_Seconds"

    ]



    - 60

)





fallback_count = int(

    fallback_rows.sum()

)





# ------------------------------------------------------------

# Negative values invalid

# ------------------------------------------------------------



df.loc[

    df["_ActualSLA_Seconds"] < 0,

    "_ActualSLA_Seconds"

] = np.nan





# ============================================================

# 19. FUNCTION:

#     SECONDS -> HH:MM:SS

#

# IMPORTANT:

# We return TEXT, not Excel time.

# This prevents Excel from showing weird formats.

# ============================================================



def seconds_to_hhmmss(seconds):



    if pd.isna(seconds):



        return ""





    total_seconds = int(

        round(float(seconds))

    )





    if total_seconds < 0:



        return ""





    hours = total_seconds // 3600





    remaining = (

        total_seconds % 3600

    )





    minutes = (

        remaining // 60

    )





    secs = (

        remaining % 60

    )





    return (

        f"{hours:02d}:"

        f"{minutes:02d}:"

        f"{secs:02d}"

    )





# ============================================================

# 20. INSTANCE-WISE SUMMARY

#

# VERY IMPORTANT:

#

# Average is calculated here:

#

# OutputPerson + Instance

#

# NOT at person level.

# ============================================================



instance_summary = (



    df.groupby(

        [

            "_OutputPerson",

            "_Instance"

        ],

        dropna=False

    )



    .agg(



        Count=(

            "_AssignedTo",

            "size"

        ),



        AverageSLA=(

            "_ActualSLA_Seconds",

            "mean"

        )



    )



    .reset_index()

)





# ============================================================

# 21. PERSON COUNT

#

# No person-level average.

# ============================================================



person_count = (



    df.groupby(

        "_OutputPerson"

    )



    .size()



    .reset_index(

        name="Count"

    )

)





# ============================================================

# 22. BUILD FINAL REPORT

# ============================================================



output_rows = []





people = sorted(



    person_count[

        "_OutputPerson"

    ].unique(),



    key=lambda x:

        str(x).lower()



)





for person in people:





# ========================================================

# PERSON HEADER

# ========================================================



    count_value = int(



        person_count[

            person_count[

                "_OutputPerson"

            ] == person

        ]["Count"].iloc[0]



    )





    output_rows.append({



        "Row Labels":

            person,



        "Count of Status":

            count_value,



        "Average of Actual SLA":

            "",



        "_Level":

            0



    })





# ========================================================

# INSTANCE ROWS

# ========================================================



    person_instances = instance_summary[

        instance_summary[

            "_OutputPerson"

        ] == person

    ].copy()





    person_instances = person_instances.sort_values(



        "_Instance",



        key=lambda x:

            x.astype(str).str.lower()



    )





    for _, row in person_instances.iterrows():





# Convert average to HH:MM:SS

        avg_sla = seconds_to_hhmmss(

            row["AverageSLA"]

        )





        output_rows.append({



            "Row Labels":

                "    "

                +

                str(

                    row["_Instance"]

                ),



            "Count of Status":

                int(

                    row["Count"]

                ),



            "Average of Actual SLA":

                avg_sla,



            "_Level":

                1



        })





# ============================================================

# 23. FINAL SUMMARY DATAFRAME

# ============================================================



summary_output = pd.DataFrame(

    output_rows

)





# ============================================================

# 24. VALIDATION

# ============================================================



worker_rows = int(



    (

        df["_AssignmentType"]

        == "WORKER"

    ).sum()



)





shift_rows = int(



    (

        df["_AssignmentType"]

        == "SHIFT"

    ).sum()



)





sla_available = int(



    df["_ActualSLA_Seconds"]

    .notna()

    .sum()



)





sla_missing = int(



    df["_ActualSLA_Seconds"]

    .isna()

    .sum()



)





validation = pd.DataFrame({



    "Metric": [



        "Original Rows",



        "Removed Blank / Zero Instance",



        "Removed Invalid Assigned To",



        "Final Included Rows",



        "Exact Name Rows -> (W)",



        "Morning/Afternoon/Evening/Night Combined Rows",



        "Shift A/B/C/D/E/F/G Combined Rows",



        "Actual SLA Available",



        "Actual SLA Missing",



        "Fallback Overall SLA - 1 Minute"



    ],



    "Count": [



        original_rows,



        removed_blank_instance,



        removed_invalid_assign,



        len(df),



        worker_rows,



        int(

            (

                df["_AssignmentType"]

                == "SHIFT"

            ).sum()

            -

            df[

                "_AssignedTo"

            ].str.lower().str.contains(

                r'\bshift\s+[a-g]\b',

                regex=True,

                na=False

            ).sum()

        ),



        int(

            df[

                "_AssignedTo"

            ].str.lower().str.contains(

                r'\bshift\s+[a-g]\b',

                regex=True,

                na=False

            ).sum()

        ),



        sla_available,



        sla_missing,



        fallback_count



    ]



})





# ============================================================

# 25. PROCESSED DATA

# ============================================================



processed_data = df.copy()





processed_data[

    "Output Person"

] = processed_data[

    "_OutputPerson"

]





processed_data[

    "Assignment Type"

] = processed_data[

    "_AssignmentType"

]





processed_data[

    "Calculated Actual SLA"

] = processed_data[

    "_ActualSLA_Seconds"

].apply(

    seconds_to_hhmmss

)





# ------------------------------------------------------------

# Remove technical columns

# ------------------------------------------------------------



technical_columns = [



    "_AssignedTo",



    "_Instance",



    "_BasePerson",



    "_AssignmentType",



    "_OutputPerson",



    "_OverallSLA_Seconds",



    "_ActualSLA_Seconds"



]





processed_data = processed_data.drop(



    columns=[

        col

        for col in technical_columns

        if col in processed_data.columns

    ]



)





# ============================================================

# 26. EXPORT

# ============================================================



output_file = (

    "SLA_Person_Instance_Summary.xlsx"

)





with pd.ExcelWriter(



    output_file,



    engine="openpyxl"



) as writer:





# --------------------------------------------------------

# Main Summary

# --------------------------------------------------------



    summary_output.drop(

        columns=["_Level"]

    ).to_excel(



        writer,



        sheet_name="SLA Summary",



        index=False



    )





# --------------------------------------------------------

# Processed Data

# --------------------------------------------------------



    processed_data.to_excel(



        writer,



        sheet_name="Processed Data",



        index=False



    )





# --------------------------------------------------------

# Validation

# --------------------------------------------------------



    validation.to_excel(



        writer,



        sheet_name="Validation",



        index=False



    )





# ============================================================

# 27. FORMAT WORKBOOK

# ============================================================



wb = load_workbook(

    output_file

)





# ============================================================

# SLA SUMMARY

# ============================================================



ws = wb[

    "SLA Summary"

]





header_fill = PatternFill(

    "solid",

    fgColor="1F4E78"

)





header_font = Font(

    color="FFFFFF",

    bold=True

)





person_fill = PatternFill(

    "solid",

    fgColor="D9EAF7"

)





# ------------------------------------------------------------

# Header formatting

# ------------------------------------------------------------



for cell in ws[1]:



    cell.fill = header_fill



    cell.font = header_font



    cell.alignment = Alignment(

        horizontal="center",

        vertical="center"

    )





# ------------------------------------------------------------

# Person rows

# ------------------------------------------------------------



for row in range(

    2,

    ws.max_row + 1

):



    label = ws.cell(

        row=row,

        column=1

    )





# Person row

    if not str(

        label.value

    ).startswith("    "):





        for col in range(

            1,

            4

        ):



            cell = ws.cell(

                row=row,

                column=col

            )



            cell.fill = person_fill



            cell.font = Font(

                bold=True

            )





# ------------------------------------------------------------

# Column widths

# ------------------------------------------------------------



ws.column_dimensions[

    "A"

].width = 42





ws.column_dimensions[

    "B"

].width = 20





ws.column_dimensions[

    "C"

].width = 25





# ------------------------------------------------------------

# Freeze header

# ------------------------------------------------------------



ws.freeze_panes = "A2"





# ============================================================

# PROCESSED DATA

# ============================================================



ws2 = wb[

    "Processed Data"

]





for cell in ws2[1]:



    cell.fill = header_fill



    cell.font = header_font





# ------------------------------------------------------------

# Auto width

# ------------------------------------------------------------



for column_cells in ws2.columns:



    max_length = 0



    column_letter = get_column_letter(

        column_cells[0].column

    )





    for cell in column_cells:



        try:



            max_length = max(

                max_length,

                len(str(cell.value))

            )



        except:



            pass





    ws2.column_dimensions[

        column_letter

    ].width = min(

        max_length + 2,

        40

    )





ws2.freeze_panes = "A2"





# ============================================================

# VALIDATION

# ============================================================



ws3 = wb[

    "Validation"

]





for cell in ws3[1]:



    cell.fill = header_fill



    cell.font = header_font





ws3.column_dimensions[

    "A"

].width = 55





ws3.column_dimensions[

    "B"

].width = 20





# ============================================================

# 28. SAVE

# ============================================================



wb.save(

    output_file

)





# ============================================================

# 29. FINAL CHECK

# ============================================================



print("\n")

print("=" * 75)

print("SUCCESS - SLA REPORT CREATED")

print("=" * 75)



print(

    f"\nOriginal rows                 : "

    f"{original_rows:,}"

)



print(

    f"Final included rows           : "

    f"{len(df):,}"

)



print(

    f"Exact Name -> (W)             : "

    f"{worker_rows:,}"

)



print(

    f"All recognized shifts         : "

    f"{shift_rows:,}"

)



print(

    f"Fallback SLA used             : "

    f"{fallback_count:,}"

)





print("\n")

print("=" * 75)

print("FINAL GROUPING LOGIC")

print("=" * 75)



print("""

EXACT NAME:



Abhinav

   ↓

Abhinav (W)





NORMAL SHIFTS:



Abhinav Morning

Abhinav Afternoon

Abhinav Evening

Abhinav Night

   ↓

Abhinav





SHIFT A-G:



Abhinav Shift A

Abhinav Shift B

Abhinav Shift C

Abhinav Shift D

Abhinav Shift E

Abhinav Shift F

Abhinav Shift G

   ↓

Abhinav





IMPORTANT:



All recognized shifts are combined.



Exact "Abhinav" stays separate as "Abhinav (W)".



Average SLA is ONLY INSTANCE-WISE.



Person row:

    Count = YES

    Average SLA = BLANK



Instance row:

    Count = YES

    Average SLA = HH:MM:SS

""")





print(

    "\nOutput file:"

)



print(

    output_file

)





# ============================================================

# 30. DOWNLOAD

# ============================================================



files.download(

    output_file

)
