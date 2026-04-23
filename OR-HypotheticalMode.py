import pandas as pd
from pulp import *

# ----------------------------
# DATASET
# ----------------------------

rooms = ["R101", "R102"]
timeslots = list(range(1, 46))  # 5 days x 9
sections = ["CS3A", "CS3B"]

subjects = {
    "Linear Algebra": 3,
    "Software Engineering": 3,
    "Operations Research": 3,
    "Automata Theory": 3,
    "Information Assurance": 3,
    "Distributed Systems": 3
}

instructors = {
    "Reynera": ["Linear Algebra"],
    "Dina": ["Linear Algebra"],
    "Nina": ["Operations Research"],
    "Noel": ["Software Engineering"],
    "Rey": ["Software Engineering", "Automata Theory"],
    "Joshua": ["Information Assurance"],
    "Simon": ["Distributed Systems"]
}

# ----------------------------
# PRE-ASSIGN INSTRUCTOR (FAST)
# ----------------------------

subject_instructor = {}

for subject in subjects:
    for instructor, teachable in instructors.items():
        if subject in teachable:
            subject_instructor[subject] = instructor
            break

# ----------------------------
# BUILD DATAFRAME (REDUCED SIZE)
# ----------------------------

rows = []

for room in rooms:
    for time in timeslots:
        for section in sections:
            for subject in subjects:
                rows.append({
                    "room": room,
                    "time": time,
                    "section": section,
                    "subject": subject,
                    "instructor": subject_instructor[subject]
                })

df = pd.DataFrame(rows)

# ----------------------------
# MODEL
# ----------------------------

model = LpProblem("Scheduling", LpMinimize)

x = {i: LpVariable(f"x_{i}", cat="Binary") for i in df.index}

# ----------------------------
# HELPERS
# ----------------------------

def get_day(t): return (t - 1) // 9
days = range(5)

# ----------------------------
# BASIC CONSTRAINTS
# ----------------------------

# Room constraint
for r in rooms:
    for t in timeslots:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "room"] == r and df.loc[i, "time"] == t
        ) <= 1

# Instructor constraint
for inst in subject_instructor.values():
    for t in timeslots:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "instructor"] == inst and df.loc[i, "time"] == t
        ) <= 1

# Section constraint
for s in sections:
    for t in timeslots:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "section"] == s and df.loc[i, "time"] == t
        ) <= 1

# Units constraint
for s in sections:
    for subj in subjects:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "section"] == s and df.loc[i, "subject"] == subj
        ) == subjects[subj]

# ----------------------------
# PATTERN ENFORCEMENT (FAST)
# ----------------------------

# pattern = 1 → (2 consecutive + 1 separate)
# pattern = 0 → (MWF-style: 3 days)
pattern = {}

day_used = {}

for s in sections:
    for subj in subjects:

        pattern[(s, subj)] = LpVariable(f"pattern_{s}_{subj}", cat="Binary")

        for d in days:
            day_used[(s, subj, d)] = LpVariable(f"day_{s}_{subj}_{d}", cat="Binary")

            slots = [t for t in timeslots if get_day(t) == d]

            model += day_used[(s, subj, d)] <= lpSum(
                x[i] for i in x
                if df.loc[i, "section"] == s
                and df.loc[i, "subject"] == subj
                and df.loc[i, "time"] in slots
            )

        total_days = lpSum(day_used[(s, subj, d)] for d in days)

        # Enforce:
        # pattern=1 → 2 days (2+1 structure)
        # pattern=0 → 3 days (MWF)
        model += total_days == 2 * pattern[(s, subj)] + 3 * (1 - pattern[(s, subj)])

# ----------------------------
# OBJECTIVE (MINIMIZE GAPS)
# ----------------------------

# Encourage compact schedules (fewer used slots)
model += lpSum(x.values())

# ----------------------------
# SOLVE (FAST SETTINGS)
# ----------------------------

model.solve(PULP_CBC_CMD(msg=1, timeLimit=30))

print("Status:", LpStatus[model.status])

# ----------------------------
# OUTPUT
# ----------------------------

days_names = ["Mon", "Tue", "Wed", "Thu", "Fri"]
time_labels = ["8-9","9-10","10-11","11-12","12-1","1-2","2-3","3-4","4-5"]

def decode(t):
    return days_names[(t-1)//9], time_labels[(t-1)%9]

schedule = []

for i in x:
    if value(x[i]) == 1:
        schedule.append(df.loc[i])

schedule_df = pd.DataFrame(schedule)

# ----------------------------
# EXCEL OUTPUT
# ----------------------------

with pd.ExcelWriter("final_schedule.xlsx", engine="openpyxl") as writer:

    for s in sections:

        grid = pd.DataFrame("", index=time_labels, columns=["M","Tu","W","Th","F"])

        for _, row in schedule_df.iterrows():
            if row["section"] == s:
                day, time_label = decode(row["time"])

                dmap = {"Mon":"M","Tue":"Tu","Wed":"W","Thu":"Th","Fri":"F"}

                grid.loc[time_label, dmap[day]] = \
                    f"{row['subject']} ({row['instructor']} - {row['room']})"

        grid.index.name = "Time"
        grid.to_excel(writer, sheet_name=s)

print("Saved to final_schedule.xlsx")