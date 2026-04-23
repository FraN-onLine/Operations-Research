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
sections = ["CS1A","IT1A","IT1B","IT1C","IT2A","IT2B","IT2C","IT3A","IT3B","IT3C","IT4A","IT4B","IT4C"]

#first semester
lecture_subjects = {
    "Computer Programming I": 2,
    "Understanding the Self": 3,
    "Linear Algebra": 3,
    "Software Engineering": 2,
    "Operations Research": 3,
    "Automata Theory": 3,
    "Distributed Systems": 2,
    "Information Assurance": 2,
    "Programming Languages": 3,
    
    #IT subjects 
    "Introduction to Computing": 2,
    "Computer Programming 1": 2,
    "Science, Technology, and Society": 3,
    "Understanding the Self": 3,
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
    
    #IT subjects
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

#FOR FIRST SEMESTER of 2026-2027
section_subjects = {
    "CS1A": ["Understanding the Self"],
    "CS4A": ["Programming Languages", ],
    "IT1A": ["Introduction to Computing", "Computer Programming 1", "Science, Technology, and Society", "Understanding the Self", "Reading Visual Art", "Movement Competency Training (MCT)", "CWTS 1/ROTC 1","Platform Technologies Lab","Computer Networks 1 Lab"],
    "IT1B": ["Introduction to Computing", "Computer Programming 1", "Science, Technology, and Society", "Understanding the Self", "Reading Visual Art", "Movement Competency Training (MCT)", "CWTS 1/ROTC 1","Platform Technologies Lab","Computer Networks 1 Lab"],
    "IT1C": ["Introduction to Computing", "Computer Programming 1", "Science, Technology, and Society", "Understanding the Self", "Reading Visual Art", "Movement Competency Training (MCT)", "CWTS 1/ROTC 1","Platform Technologies Lab","Computer Networks 1 Lab"],
    "IT2A": ["Database Systems", "Operating System", "Data Structures & Algorithms", "Introduction to Game Development", "Ethics", "Fundamentals of Accounting for IT", "Environmental Science", "Dance","Database Systems Lab","Operating System Lab","Data Structures Lab","Introduction to Game Development Lab"],
    "IT2B": ["Database Systems", "Operating System", "Data Structures & Algorithms", "Introduction to Game Development", "Ethics", "Fundamentals of Accounting for IT", "Environmental Science", "Dance","Database Systems Lab","Operating System Lab","Data Structures Lab","Introduction to Game Development Lab"],
    "IT2C": ["Database Systems", "Operating System", "Data Structures & Algorithms", "Introduction to Game Development", "Ethics", "Fundamentals of Accounting for IT", "Environmental Science", "Dance","Database Systems Lab","Operating System Lab","Data Structures Lab","Introduction to Game Development Lab"],
    "IT3A": ["Computer Networks 1", "Platform Technologies", "Data Analytics", "Information and Project Management", "System Administration & Maintenance", "Fundamentals of Business Analytics", "Life and Works of Rizal","Computer Networks 1 Lab", "Platform Technologies Lab", "Data Analytics Lab", "Information and Project Management Lab", "System Administration & Maintenance Lab", "Fundamentals of Business Analytics Lab"],
    "IT3B": ["Computer Networks 1", "Platform Technologies", "Data Analytics", "Information and Project Management", "System Administration & Maintenance", "Fundamentals of Business Analytics", "Life and Works of Rizal","Computer Networks 1 Lab", "Platform Technologies Lab", "Data Analytics Lab", "Information and Project Management Lab", "System Administration & Maintenance Lab", "Fundamentals of Business Analytics Lab"],
    "IT3C": ["Computer Networks 1", "Platform Technologies", "Data Analytics", "Information and Project Management", "System Administration & Maintenance", "Fundamentals of Business Analytics", "Life and Works of Rizal","Computer Networks 1 Lab", "Platform Technologies Lab", "Data Analytics Lab", "Information and Project Management Lab", "System Administration & Maintenance Lab", "Fundamentals of Business Analytics Lab"],
    "IT4A": ["Social Issues & Ethics in Computing", "Information Assurance & Security 2", "Multimedia Systems", "IT Seminar", "Capstone Project Writing","Information Assurance & Security 2 Lab", "Multimedia Systems Lab", "IT Seminar Lab", "Capstone Project Writing Lab"],
    "IT4B": ["Social Issues & Ethics in Computing", "Information Assurance & Security 2", "Multimedia Systems", "IT Seminar", "Capstone Project Writing","Information Assurance & Security 2 Lab", "Multimedia Systems Lab", "IT Seminar Lab", "Capstone Project Writing Lab"],
    "IT4C": ["Social Issues & Ethics in Computing", "Information Assurance & Security 2", "Multimedia Systems", "IT Seminar", "Capstone Project Writing","Information Assurance & Security 2 Lab", "Multimedia Systems Lab", "IT Seminar Lab", "Capstone Project Writing Lab"],
}

# Precompute valid lab starts (same-day guarantee)
valid_lab_starts = [1,4,7,10,13,16,19,22,25,28,31,34,37,40,43]

lunch_slots = [5, 14, 23, 32, 41]

# ----------------------------
# MODEL
# ----------------------------

model = LpProblem("Scheduling", LpMinimize)

# ----------------------------
# VARIABLES
# ----------------------------

x = LpVariable.dicts("Lecture",
    [(r,t,sec,sub)
     for r in lecture_rooms
     for t in timeslots
     for sec in sections
     for sub in lecture_subjects
     if sub in section_subjects[sec]],
    cat="Binary"
)

z = LpVariable.dicts("Lab",
    [(r,t,sec,sub)
     for r in lab_rooms
     for t in valid_lab_starts
     for sec in sections
     for sub in lab_subjects
     if sub in section_subjects[sec]],
    cat="Binary"
)

y = LpVariable.dicts("Occupied",
    [(sec,t) for sec in sections for t in timeslots],
    cat="Binary"
)

# ----------------------------
# INDEXING (OPTIMIZED)
# ----------------------------

x_sec_t = defaultdict(list)
x_room_t = defaultdict(list)
x_sub_t = defaultdict(list)
x_sec_sub = defaultdict(list)

for (r,t,sec,sub) in x:
    x_sec_t[(sec,t)].append(x[(r,t,sec,sub)])
    x_room_t[(r,t)].append(x[(r,t,sec,sub)])
    x_sub_t[(sub,t)].append(x[(r,t,sec,sub)])
    x_sec_sub[(sec,sub)].append(x[(r,t,sec,sub)])

z_sec_t = defaultdict(list)
z_room_t = defaultdict(list)
z_sub_t = defaultdict(list)
z_sec_sub = defaultdict(list)

for (r,t,sec,sub) in z:
    for dt in range(3):
        z_sec_t[(sec,t+dt)].append(z[(r,t,sec,sub)])
        z_room_t[(r,t+dt)].append(z[(r,t,sec,sub)])
        z_sub_t[(sub,t+dt)].append(z[(r,t,sec,sub)])
    z_sec_sub[(sec,sub)].append(z[(r,t,sec,sub)])

# ----------------------------
# CONSTRAINTS
# ----------------------------

# Link occupied slots
for sec in sections:
    for t in timeslots:
        model += y[(sec,t)] == (
            lpSum(x_sec_t[(sec,t)]) +
            lpSum(z_sec_t[(sec,t)])
        )

# Room capacity
for r in lecture_rooms:
    for t in timeslots:
        model += lpSum(x_room_t[(r,t)]) <= 1

for r in lab_rooms:
    for t in timeslots:
        model += lpSum(z_room_t[(r,t)]) <= 1

# Section cannot overlap
for sec in sections:
    for t in timeslots:
        model += (
            lpSum(x_sec_t[(sec,t)]) +
            lpSum(z_sec_t[(sec,t)])
        ) <= 1

# No 2 sections same subject same time
for sub in lecture_subjects:
    for t in timeslots:
        model += lpSum(x_sub_t[(sub,t)]) <= 1

for sub in lab_subjects:
    for t in timeslots:
        model += lpSum(z_sub_t[(sub,t)]) <= 1

# Required units
for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            model += lpSum(x_sec_sub[(sec,sub)]) == lecture_subjects[sub]

    for sub in lab_subjects:
        if sub in section_subjects[sec]:
            model += lpSum(z_sec_sub[(sec,sub)]) == lab_subjects[sub]

# Lunch break enforced
for sec in sections:
    for t in lunch_slots:
        model += y[(sec,t)] == 0

# ----------------------------
# OBJECTIVE FUNCTION
# ----------------------------

# 1. Minimize vacant time
vacant_penalty = lpSum(
    1 - y[(sec,t)]
    for sec in sections for t in timeslots
)

# 2. Balance daily load
balance_penalty = []
avg = sum(lecture_subjects.values()) / 5

for sec in sections:
    for d in range(5):
        day_slots = range(1 + d*9, 10 + d*9)
        load = lpSum(y[(sec,t)] for t in day_slots)

        p = LpVariable(f"balance_{sec}_{d}", lowBound=0)
        model += p >= load - avg
        model += p >= avg - load
        balance_penalty.append(p)

# 3. Same timeslot across days
time_penalty = []

for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            for d in range(4):
                for k in range(1,10):
                    if k == 5: continue  # skip lunch

                    t1 = d*9 + k
                    t2 = (d+1)*9 + k

                    p = LpVariable(f"time_{sec}_{sub}_{d}_{k}", lowBound=0)

                    model += p >= (
                        lpSum(x[(r,t1,sec,sub)] for r in lecture_rooms) -
                        lpSum(x[(r,t2,sec,sub)] for r in lecture_rooms)
                    )
                    model += p >= (
                        lpSum(x[(r,t2,sec,sub)] for r in lecture_rooms) -
                        lpSum(x[(r,t1,sec,sub)] for r in lecture_rooms)
                    )

                    time_penalty.append(p)

# 4. Room switching penalty
room_penalty = []

for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            for r1 in lecture_rooms:
                for r2 in lecture_rooms:
                    if r1 < r2:
                        p = LpVariable(f"room_{sec}_{sub}_{r1}_{r2}", lowBound=0)

                        model += p >= (
                            lpSum(x[(r1,t,sec,sub)] for t in timeslots) +
                            lpSum(x[(r2,t,sec,sub)] for t in timeslots) - 1
                        )

                        room_penalty.append(p)

# FINAL OBJECTIVE
model += (
    20 * vacant_penalty +
    10 * lpSum(balance_penalty) +
    5 * lpSum(time_penalty) +
    8 * lpSum(room_penalty)
)

# ----------------------------
# SOLVE
# ----------------------------

print("Solving...")
model.solve()
print("Status:", LpStatus[model.status])

# ----------------------------
# BUILD SCHEDULE
# ----------------------------

schedule = []

# lectures
for (r,t,sec,sub) in x:
    if value(x[(r,t,sec,sub)]) == 1:
        schedule.append([sec, sub, r, t, "Lecture"])

# labs (expand 3 hours)
for (r,t,sec,sub) in z:
    if value(z[(r,t,sec,sub)]) == 1:
        for dt in range(3):
            schedule.append([sec, sub, r, t+dt, "Lab"])

schedule_df = pd.DataFrame(
    schedule,
    columns=["Section","Subject","Room","Time","Type"]
)

# ----------------------------
# TIME DECODER
# ----------------------------

days = ["Mon","Tue","Wed","Thu","Fri"]
times = ["8-9","9-10","10-11","11-12","12-1","1-2","2-3","3-4","4-5"]

def decode(t):
    return days[(t-1)//9], times[(t-1)%9]

# ----------------------------
# EXCEL OUTPUT
# ----------------------------

with pd.ExcelWriter("Full_Timetable.xlsx", engine="openpyxl") as writer:

    for sec in sections:
        grid = {time: {day:"" for day in days} for time in times}
        type_map = {}

        for _, row in schedule_df.iterrows():
            if row["Section"] == sec:
                d, t = decode(row["Time"])
                grid[t][d] = f"{row['Subject']} ({row['Room']})"
                type_map[(t,d)] = row["Type"]

        df_grid = pd.DataFrame(grid).T
        df_grid.to_excel(writer, sheet_name=sec)

        ws = writer.sheets[sec]

        # Auto width
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 30

        lecture_fill = PatternFill(start_color="ADD8E6", fill_type="solid")
        lab_fill = PatternFill(start_color="90EE90", fill_type="solid")

        for i, time in enumerate(times, start=2):
            for j, day in enumerate(days, start=2):
                cell = ws.cell(row=i, column=j)

                if (time,day) in type_map:
                    if type_map[(time,day)] == "Lecture":
                        cell.fill = lecture_fill
                    else:
                        cell.fill = lab_fill

print("✅ Saved: Full_Timetable.xlsx")