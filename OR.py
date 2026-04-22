import pandas as pd
from pulp import *

# ----------------------------
# DATASET GENERATION
# ----------------------------

rooms = ["R101", "R102", "R103"]
timeslots = list(range(1, 46))  # 45 slots (5 days x 9)
sections = ["CS3A", "CS3B"]

subjects = {
    "Linear Algebra": 3,
    "Software Engineering": 3,
    "Operations Research": 3,
    "Automata Theory": 3
}

instructors = {
    "Reynera": ["Linear Algebra"],
    "Dina": ["Linear Algebra"],
    "Nina": ["Operations Research"],
    "Noel": ["Software Engineering"],
    "Rey": ["Software Engineering", "Automata Theory"]
}

section_subjects = {
    "CS3A": list(subjects.keys()),
    "CS3B": list(subjects.keys())
}

rows = []

for room in rooms:
    for time in timeslots:
        for section in sections:
            for subject in section_subjects[section]:
                for instructor, teachable in instructors.items():
                    if subject in teachable:
                        rows.append({
                            "room": room,
                            "time": time,
                            "section": section,
                            "subject": subject,
                            "instructor": instructor,
                            "units": subjects[subject]
                        })

df = pd.DataFrame(rows)

# ----------------------------
# MODEL
# ----------------------------

model = LpProblem("Class_Scheduling", LpMinimize)

# Decision Variables
x = {i: LpVariable(f"x_{i}", cat="Binary") for i in df.index}

# ----------------------------
# HELPER VARIABLES (y)
# ----------------------------
y = {}

for section in sections:
    for time in timeslots:
        y[(section, time)] = LpVariable(f"y_{section}_{time}", cat="Binary")

        model += y[(section, time)] == lpSum(
            x[i] for i in x
            if df.loc[i, "section"] == section and df.loc[i, "time"] == time
        )

# ----------------------------
# INSTRUCTOR CONSISTENCY (z)
# ----------------------------
z = {}

for section in sections:
    for subject in subjects:
        for instructor in instructors:
            if subject in instructors[instructor]:
                z[(section, subject, instructor)] = LpVariable(
                    f"z_{section}_{subject}_{instructor}", cat="Binary"
                )

# Exactly ONE instructor per subject per section
for section in sections:
    for subject in subjects:
        model += lpSum(
            z[(section, subject, instructor)]
            for instructor in instructors
            if (section, subject, instructor) in z
        ) == 1

# Link x and z
for i in df.index:
    row = df.loc[i]
    key = (row["section"], row["subject"], row["instructor"])
    if key in z:
        model += x[i] <= z[key]

# ----------------------------
# CONSTRAINTS
# ----------------------------

# A. Room constraint
for room in rooms:
    for time in timeslots:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "room"] == room and df.loc[i, "time"] == time
        ) <= 1

# B. Instructor constraint
for instructor in instructors:
    for time in timeslots:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "instructor"] == instructor and df.loc[i, "time"] == time
        ) <= 1

# C. Section constraint
for section in sections:
    for time in timeslots:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "section"] == section and df.loc[i, "time"] == time
        ) <= 1

# D. Unit constraint
for section in sections:
    for subject in subjects:
        model += lpSum(
            x[i] for i in x
            if df.loc[i, "section"] == section and df.loc[i, "subject"] == subject
        ) == subjects[subject]

# ----------------------------
# PENALTIES & REWARDS
# ----------------------------

# Vacant slots penalty
vacant_penalty = lpSum(
    (1 - y[(section, time)])
    for section in sections
    for time in timeslots
)

# Too many consecutive (>5)
consecutive_penalties = []

for section in sections:
    for i in range(len(timeslots) - 5):
        window = timeslots[i:i+6]

        p = LpVariable(f"consec_{section}_{i}", lowBound=0)

        model += p >= (
            lpSum(y[(section, t)] for t in window) - 5
        )

        consecutive_penalties.append(p)

# Reward consecutive same-subject sessions
consecutive_rewards = []

for section in sections:
    for subject in subjects:
        for t in timeslots[:-1]:
            pair = LpVariable(f"pair_{section}_{subject}_{t}", cat="Binary")

            model += pair <= lpSum(
                x[i] for i in x
                if df.loc[i, "section"] == section
                and df.loc[i, "subject"] == subject
                and df.loc[i, "time"] == t
            )

            model += pair <= lpSum(
                x[i] for i in x
                if df.loc[i, "section"] == section
                and df.loc[i, "subject"] == subject
                and df.loc[i, "time"] == t + 1
            )

            consecutive_rewards.append(pair)

# ----------------------------
# OBJECTIVE
# ----------------------------

model += (
    15 * vacant_penalty
    + 10 * lpSum(consecutive_penalties)
    - 20 * lpSum(consecutive_rewards)
)

# ----------------------------
# SOLVE
# ----------------------------

print("Solving...")
model.solve()
print("Status:", LpStatus[model.status])

# ----------------------------
# OUTPUT
# ----------------------------

days = ["Mon", "Tue", "Wed", "Thu", "Fri"]

time_labels = [
    "8-9", "9-10", "10-11", "11-12",
    "12-1", "1-2", "2-3", "3-4", "4-5"
]

def decode_timeslot(t):
    day_index = (t - 1) // 9
    slot_index = (t - 1) % 9
    return days[day_index], time_labels[slot_index]

schedule = []

for i in x:
    if value(x[i]) == 1:
        schedule.append(df.loc[i])

schedule_df = pd.DataFrame(schedule)

# ----------------------------
# EXCEL OUTPUT (5x9 per section)
# ----------------------------

with pd.ExcelWriter("final_schedule.xlsx", engine="openpyxl") as writer:

    for section in sections:

        grid = pd.DataFrame(
            "",
            index=time_labels,
            columns=["M", "Tu", "W", "Th", "F"]
        )

        for _, row in schedule_df.iterrows():
            if row["section"] == section:
                day, time_label = decode_timeslot(row["time"])

                day_map = {
                    "Mon": "M",
                    "Tue": "Tu",
                    "Wed": "W",
                    "Thu": "Th",
                    "Fri": "F"
                }

                d = day_map[day]

                grid.loc[time_label, d] = \
                    f"{row['subject']} ({row['instructor']})"

        grid.index.name = "Time"
        grid.to_excel(writer, sheet_name=section)

print("Saved to final_schedule.xlsx")