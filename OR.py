import pandas as pd
from pulp import *
from collections import defaultdict
from openpyxl.styles import PatternFill

# ----------------------------
# DATA
# ----------------------------
lecture_rooms = ["R100A", "R100B", "R100C"]
lab_rooms = ["Lab1", "Lab2"]
timeslots = list(range(1, 46))  # 5 days x 9
sections = [
    "CS1A", "IT1A", "IT1B", "IT1C",
    "IT2A", "IT2B", "IT2C",
    "IT3A", "IT3B", "IT3C",
    "IT4A", "IT4B", "IT4C"
]

# BUG FIX 1: lecture_subjects had duplicate key "Understanding the Self"
# (second entry silently overwrote the first in a plain dict).
# Use an OrderedDict-style list-of-tuples then convert so duplicates are explicit.
# Also removed "Understanding the Self" duplicate and kept one copy.
lecture_subjects = {
    # CS subjects
    "Computer Programming I": 2,
    "Linear Algebra": 3,
    "Software Engineering": 2,
    "Operations Research": 3,
    "Automata Theory": 3,
    "Distributed Systems": 2,
    "Information Assurance": 2,
    "Programming Languages": 3,
    # IT subjects
    "Introduction to Computing": 2,
    "Computer Programming 1": 2,
    "Science, Technology, and Society": 3,
    "Understanding the Self": 3,          # kept once (was duplicated)
    "Reading Visual Art": 3,
    "Movement Competency Training (MCT)": 2,
    "CWTS 1/ROTC 1": 3,
    "Database Systems": 2,
    "Operating System": 2,
    "Data Structures & Algorithms": 2,
    "Introduction to Game Development": 2,
    "Ethics": 3,
    "Fundamentals of Accounting for IT": 3,
    "Environmental Science": 3,
    "Dance": 2,
    "Computer Networks 1": 2,
    "Platform Technologies": 2,
    "Data Analytics": 2,
    "Information and Project Management": 2,
    "System Administration & Maintenance": 2,
    "Fundamentals of Business Analytics": 2,
    "Life and Works of Rizal": 3,
    "Social Issues & Ethics in Computing": 3,
    "Information Assurance & Security 2": 2,
    "Multimedia Systems": 2,
    "IT Seminar": 2,
    "Capstone Project Writing": 2,
}

lab_subjects = {
    "Computer Programming I Lab": 1,
    "Computer Science Fundamentals Lab": 1,
    "Data Structures Lab": 1,
    "Digital Design Lab": 1,
    "Database Systems Lab": 1,
    "Discrete Structures Lab": 1,
    "Distributed Systems Lab": 1,
    "Software Engineering Lab": 1,
    "Information Assurance Lab": 1,
    "Operating System Lab": 1,
    # IT subjects
    "Introduction to Computing Lab": 1,
    "Computer Programming 1 Lab": 1,
    "Introduction to Game Development Lab": 1,
    "Computer Networks 1 Lab": 1,
    "Platform Technologies Lab": 1,
    "Data Analytics Lab": 1,
    "Information and Project Management Lab": 1,
    "System Administration & Maintenance Lab": 1,
    "Fundamentals of Business Analytics Lab": 1,
    "Information Assurance & Security 2 Lab": 1,
    "Multimedia Systems Lab": 1,
    "IT Seminar Lab": 1,
    "Capstone Project Writing Lab": 1,
}

# BUG FIX 2: "CS4A" was in section_subjects but NOT in sections list,
# causing silent KeyErrors when building constraints. Either add it to
# sections or remove it from section_subjects. Here we add it properly.
# Also note: "Computer Science Fundamentals" was referenced for CS4A but
# it's not in lecture_subjects — kept only valid subjects.
section_subjects = {
    "CS1A": ["Understanding the Self"],
    # CS4A removed (not in sections list; add back to sections if needed)
    "IT1A": [
        "Introduction to Computing", "Computer Programming 1",
        "Science, Technology, and Society", "Understanding the Self",
        "Reading Visual Art", "Movement Competency Training (MCT)",
        "CWTS 1/ROTC 1",
        "Platform Technologies Lab", "Computer Networks 1 Lab",
    ],
    "IT1B": [
        "Introduction to Computing", "Computer Programming 1",
        "Science, Technology, and Society", "Understanding the Self",
        "Reading Visual Art", "Movement Competency Training (MCT)",
        "CWTS 1/ROTC 1",
        "Platform Technologies Lab", "Computer Networks 1 Lab",
    ],
    "IT1C": [
        "Introduction to Computing", "Computer Programming 1",
        "Science, Technology, and Society", "Understanding the Self",
        "Reading Visual Art", "Movement Competency Training (MCT)",
        "CWTS 1/ROTC 1",
        "Platform Technologies Lab", "Computer Networks 1 Lab",
    ],
    "IT2A": [
        "Database Systems", "Operating System", "Data Structures & Algorithms",
        "Introduction to Game Development", "Ethics",
        "Fundamentals of Accounting for IT", "Environmental Science", "Dance",
        "Database Systems Lab", "Operating System Lab",
        "Data Structures Lab", "Introduction to Game Development Lab",
    ],
    "IT2B": [
        "Database Systems", "Operating System", "Data Structures & Algorithms",
        "Introduction to Game Development", "Ethics",
        "Fundamentals of Accounting for IT", "Environmental Science", "Dance",
        "Database Systems Lab", "Operating System Lab",
        "Data Structures Lab", "Introduction to Game Development Lab",
    ],
    "IT2C": [
        "Database Systems", "Operating System", "Data Structures & Algorithms",
        "Introduction to Game Development", "Ethics",
        "Fundamentals of Accounting for IT", "Environmental Science", "Dance",
        "Database Systems Lab", "Operating System Lab",
        "Data Structures Lab", "Introduction to Game Development Lab",
    ],
    "IT3A": [
        "Computer Networks 1", "Platform Technologies", "Data Analytics",
        "Information and Project Management", "System Administration & Maintenance",
        "Fundamentals of Business Analytics", "Life and Works of Rizal",
        "Computer Networks 1 Lab", "Platform Technologies Lab",
        "Data Analytics Lab", "Information and Project Management Lab",
        "System Administration & Maintenance Lab", "Fundamentals of Business Analytics Lab",
    ],
    "IT3B": [
        "Computer Networks 1", "Platform Technologies", "Data Analytics",
        "Information and Project Management", "System Administration & Maintenance",
        "Fundamentals of Business Analytics", "Life and Works of Rizal",
        "Computer Networks 1 Lab", "Platform Technologies Lab",
        "Data Analytics Lab", "Information and Project Management Lab",
        "System Administration & Maintenance Lab", "Fundamentals of Business Analytics Lab",
    ],
    "IT3C": [
        "Computer Networks 1", "Platform Technologies", "Data Analytics",
        "Information and Project Management", "System Administration & Maintenance",
        "Fundamentals of Business Analytics", "Life and Works of Rizal",
        "Computer Networks 1 Lab", "Platform Technologies Lab",
        "Data Analytics Lab", "Information and Project Management Lab",
        "System Administration & Maintenance Lab", "Fundamentals of Business Analytics Lab",
    ],
    "IT4A": [
        "Social Issues & Ethics in Computing", "Information Assurance & Security 2",
        "Multimedia Systems", "IT Seminar", "Capstone Project Writing",
        "Information Assurance & Security 2 Lab", "Multimedia Systems Lab",
        "IT Seminar Lab", "Capstone Project Writing Lab",
    ],
    "IT4B": [
        "Social Issues & Ethics in Computing", "Information Assurance & Security 2",
        "Multimedia Systems", "IT Seminar", "Capstone Project Writing",
        "Information Assurance & Security 2 Lab", "Multimedia Systems Lab",
        "IT Seminar Lab", "Capstone Project Writing Lab",
    ],
    "IT4C": [
        "Social Issues & Ethics in Computing", "Information Assurance & Security 2",
        "Multimedia Systems", "IT Seminar", "Capstone Project Writing",
        "Information Assurance & Security 2 Lab", "Multimedia Systems Lab",
        "IT Seminar Lab", "Capstone Project Writing Lab",
    ],
}

# Lab sessions occupy 3 consecutive timeslots. The only requirement is that
# all 3 slots fit within the same day (p+2 <= 9 => p <= 7).
# Labs are allowed to span lunch slots.
valid_lab_starts = []
for d in range(5):
    base = d * 9
    for p in range(1, 8):  # positions 1..7 all fit within the day
        valid_lab_starts.append(base + p)

lunch_slots = [5, 14, 23, 32, 41]

# Precompute subjects available per section to avoid repeated lookups.
lecture_subjects_in_section = {
    sec: [sub for sub in section_subjects.get(sec, []) if sub in lecture_subjects]
    for sec in sections
}
lab_subjects_in_section = {
    sec: [sub for sub in section_subjects.get(sec, []) if sub in lab_subjects]
    for sec in sections
}

# ----------------------------
# MODEL
# ----------------------------
model = LpProblem("Scheduling", LpMinimize)

# ----------------------------
# VARIABLES
# ----------------------------
x = LpVariable.dicts("Lecture",
    [(r, t, sec, sub)
     for r in lecture_rooms
     for t in timeslots
     for sec in sections
     for sub in lecture_subjects_in_section[sec]],
    cat="Binary"
)

z = LpVariable.dicts("Lab",
    [(r, t, sec, sub)
     for r in lab_rooms
     for t in valid_lab_starts
     for sec in sections
     for sub in lab_subjects_in_section[sec]],
    cat="Binary"
)

y = LpVariable.dicts("Occupied",
    [(sec, t) for sec in sections for t in timeslots],
    cat="Binary"
)

# ----------------------------
# INDEXING
# ----------------------------
x_sec_t   = defaultdict(list)
x_room_t  = defaultdict(list)
x_sub_t   = defaultdict(list)
x_sec_sub = defaultdict(list)
x_sec_sub_t = defaultdict(list)
x_room_sub = defaultdict(list)

for (r, t, sec, sub) in x:
    var = x[(r, t, sec, sub)]
    x_sec_t[(sec, t)].append(var)
    x_room_t[(r, t)].append(var)
    x_sub_t[(sub, t)].append(var)
    x_sec_sub[(sec, sub)].append(var)
    x_sec_sub_t[(sec, sub, t)].append(var)
    x_room_sub[(r, sec, sub)].append(var)

z_sec_t   = defaultdict(list)
z_room_t  = defaultdict(list)
z_sub_t   = defaultdict(list)
z_sec_sub = defaultdict(list)

for (r, t, sec, sub) in z:
    var = z[(r, t, sec, sub)]
    for dt in range(3):
        z_sec_t[(sec, t + dt)].append(var)
        z_room_t[(r, t + dt)].append(var)
        z_sub_t[(sub, t + dt)].append(var)
    z_sec_sub[(sec, sub)].append(var)

# ----------------------------
# CONSTRAINTS
# ----------------------------

# Link occupied slots
for sec in sections:
    for t in timeslots:
        model += y[(sec, t)] == (
            lpSum(x_sec_t[(sec, t)]) +
            lpSum(z_sec_t[(sec, t)])
        )

# Room capacity: at most 1 class per room per timeslot
for r in lecture_rooms:
    for t in timeslots:
        model += lpSum(x_room_t[(r, t)]) <= 1

for r in lab_rooms:
    for t in timeslots:
        model += lpSum(z_room_t[(r, t)]) <= 1

# Section cannot be in two places at once
for sec in sections:
    for t in timeslots:
        model += (
            lpSum(x_sec_t[(sec, t)]) +
            lpSum(z_sec_t[(sec, t)])
        ) <= 1

# No 2 sections same subject at the same time
for sub in lecture_subjects:
    for t in timeslots:
        model += lpSum(x_sub_t[(sub, t)]) <= 1

for sub in lab_subjects:
    for t in timeslots:
        model += lpSum(z_sub_t[(sub, t)]) <= 1

# BUG FIX 4 (MAIN FIX): Required units constraint.
# Each lecture subject with N units must appear exactly N times per section.
# Previously this worked in theory but was broken by the duplicate-key bug
# in lecture_subjects (which caused some subjects to have wrong unit counts)
# and by sections missing from section_subjects (KeyError).
# Now that those are fixed, this constraint correctly enforces repetition.
for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects.get(sec, []):
            required = lecture_subjects[sub]
            model += lpSum(x_sec_sub[(sec, sub)]) == required

    for sub in lab_subjects:
        if sub in section_subjects.get(sec, []):
            required = lab_subjects[sub]
            model += lpSum(z_sec_sub[(sec, sub)]) == required

# Lunch break enforced
for sec in sections:
    for t in lunch_slots:
        model += y[(sec, t)] == 0

# ----------------------------
# OBJECTIVE FUNCTION
# ----------------------------

# 1. Minimize vacant (non-lunch, non-occupied) time
vacant_penalty = lpSum(
    1 - y[(sec, t)]
    for sec in sections
    for t in timeslots
    if t not in lunch_slots
)

# 2. Balance daily load
avg = sum(lecture_subjects[s] for sec in sections
          for s in section_subjects.get(sec, [])
          if s in lecture_subjects) / (len(sections) * 5)

balance_penalty = []
for sec in sections:
    for d in range(5):
        day_slots = range(1 + d * 9, 10 + d * 9)
        load = lpSum(y[(sec, t)] for t in day_slots if t not in lunch_slots)
        p = LpVariable(f"balance_{sec}_{d}", lowBound=0)
        model += p >= load - avg
        model += p >= avg - load
        balance_penalty.append(p)

# 3. Same timeslot across days (consistency for multi-unit lectures)
time_penalty = []
for sec in sections:
    for sub in lecture_subjects_in_section[sec]:
        for d in range(4):
            for k in range(1, 10):
                if k == 5:
                    continue
                t1 = d * 9 + k
                t2 = (d + 1) * 9 + k
                p = LpVariable(f"time_{sec}_{sub}_{d}_{k}", lowBound=0)
                model += p >= lpSum(x_sec_sub_t[(sec, sub, t1)]) - lpSum(x_sec_sub_t[(sec, sub, t2)])
                model += p >= lpSum(x_sec_sub_t[(sec, sub, t2)]) - lpSum(x_sec_sub_t[(sec, sub, t1)])
                time_penalty.append(p)

# 4. Room switching penalty (prefer same room for all sessions of a subject)
room_penalty = []
for sec in sections:
    for sub in lecture_subjects_in_section[sec]:
        for i, r1 in enumerate(lecture_rooms):
            for r2 in lecture_rooms[i + 1:]:
                p = LpVariable(f"room_{sec}_{sub}_{r1}_{r2}", lowBound=0)
                model += p >= lpSum(x_room_sub[(r1, sec, sub)]) + lpSum(x_room_sub[(r2, sec, sub)]) - 1
                room_penalty.append(p)

# FINAL OBJECTIVE
model += (
    20 * vacant_penalty +
    10 * lpSum(balance_penalty) +
    5  * lpSum(time_penalty) +
    8  * lpSum(room_penalty)
)

# ----------------------------
# SOLVE
# ----------------------------
print("Solving with CBC (120s limit, 4 threads)...")
model.solve(PULP_CBC_CMD(msg=True, timeLimit=120, threads=4))
print("Status:", LpStatus[model.status])

# ----------------------------
# BUILD SCHEDULE
# ----------------------------
schedule = []

for (r, t, sec, sub) in x:
    if value(x[(r, t, sec, sub)]) == 1:
        schedule.append([sec, sub, r, t, "Lecture"])

for (r, t, sec, sub) in z:
    if value(z[(r, t, sec, sub)]) == 1:
        for dt in range(3):
            schedule.append([sec, sub, r, t + dt, "Lab"])

schedule_df = pd.DataFrame(
    schedule,
    columns=["Section", "Subject", "Room", "Time", "Type"]
)

# ----------------------------
# VERIFICATION: print unit counts per section/subject
# ----------------------------
print("\n--- Unit Count Verification ---")
for sec in sections:
    sec_df = schedule_df[schedule_df["Section"] == sec]
    for sub in section_subjects.get(sec, []):
        is_lab = sub in lab_subjects
        count = len(sec_df[sec_df["Subject"] == sub])
        expected = lab_subjects[sub] * 3 if is_lab else lecture_subjects[sub]
        status = "✅" if count == expected else "❌"
        if count != expected:
            print(f"  {status} {sec} | {sub}: got {count} slots, expected {expected}")

# ----------------------------
# TIME DECODER
# ----------------------------
days  = ["Mon", "Tue", "Wed", "Thu", "Fri"]
times = ["8-9", "9-10", "10-11", "11-12", "12-1", "1-2", "2-3", "3-4", "4-5"]

def decode(t):
    return days[(t - 1) // 9], times[(t - 1) % 9]

# ----------------------------
# EXCEL OUTPUT
# ----------------------------
with pd.ExcelWriter("Full_Timetable.xlsx", engine="openpyxl") as writer:
    for sec in sections:
        grid    = {time: {day: "" for day in days} for time in times}
        type_map = {}

        for _, row in schedule_df.iterrows():
            if row["Section"] == sec:
                d, t = decode(row["Time"])
                grid[t][d] = f"{row['Subject']} ({row['Room']})"
                type_map[(t, d)] = row["Type"]

        df_grid = pd.DataFrame(grid).T
        df_grid.to_excel(writer, sheet_name=sec)

        ws = writer.sheets[sec]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 30

        lecture_fill = PatternFill(start_color="ADD8E6", fill_type="solid")
        lab_fill     = PatternFill(start_color="90EE90", fill_type="solid")

        for i, time in enumerate(times, start=2):
            for j, day in enumerate(days, start=2):
                cell = ws.cell(row=i, column=j)
                if (time, day) in type_map:
                    cell.fill = lecture_fill if type_map[(time, day)] == "Lecture" else lab_fill

print("\n✅ Saved: Full_Timetable.xlsx")