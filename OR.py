import pandas as pd
from pulp import *

# ----------------------------
# DATA
# ----------------------------

lecture_rooms = ["R101", "R102"]
lab_rooms = ["Lab1", "Lab2"]

timeslots = list(range(1, 46))  # 5 days x 9
sections = ["CS3A", "CS3B"]

lecture_subjects = {
    "Linear Algebra": 3,
    "Software Engineering": 2,
    "Operations Research": 3,
    "Automata Theory": 3,
    "Distributed Systems": 2,
    "Information Assurance": 2
}

lab_subjects = {
    "Distributed Systems Lab": 1,
    "Software Engineering Lab": 1,
    "Information Assurance Lab": 1
}

section_subjects = {
    "CS3A": list(lecture_subjects.keys()) + list(lab_subjects.keys()),
    "CS3B": list(lecture_subjects.keys()) + list(lab_subjects.keys())
}

# ----------------------------
# MODEL
# ----------------------------

model = LpProblem("Scheduling", LpMinimize)

# ----------------------------
# VARIABLES
# ----------------------------

# Lecture variables
x = {}
for r in lecture_rooms:
    for t in timeslots:
        for sec in sections:
            for sub in lecture_subjects:
                if sub in section_subjects[sec]:
                    x[(r,t,sec,sub)] = LpVariable(f"x_{r}_{t}_{sec}_{sub}", cat="Binary")

# Lab variables (start time only)
z = {}
for r in lab_rooms:
    for t in timeslots:
        if (t - 1) % 9 <= 6:
            for sec in sections:
                for sub in lab_subjects:
                    if sub in section_subjects[sec]:
                        z[(r,t,sec,sub)] = LpVariable(f"z_{r}_{t}_{sec}_{sub}", cat="Binary")

# Section-time indicator
y = {}
for sec in sections:
    for t in timeslots:
        y[(sec,t)] = LpVariable(f"y_{sec}_{t}", cat="Binary")

# ----------------------------
# CONSTRAINTS
# ----------------------------

# Link y with x and z
for sec in sections:
    for t in timeslots:
        model += y[(sec,t)] == (
            lpSum(x[(r,t,sec,s)] for (r,t2,sec2,s) in x if sec2==sec and t2==t)
            +
            lpSum(z[(r,start,sec,s)] for (r,start,sec2,s) in z if sec2==sec and start<=t<=start+2)
        )

# Lecture room constraint
for r in lecture_rooms:
    for t in timeslots:
        model += lpSum(x[(r,t,sec,sub)]
                       for (r2,t2,sec,sub) in x if r2==r and t2==t) <= 1

# Lab room constraint
for r in lab_rooms:
    for t in timeslots:
        model += lpSum(z[(r,start,sec,sub)]
                       for (r2,start,sec,sub) in z if r2==r and start<=t<=start+2) <= 1

# Section constraint
for sec in sections:
    for t in timeslots:
        model += (
            lpSum(x[(r,t,sec,s)] for (r,t2,sec2,s) in x if sec2==sec and t2==t)
            +
            lpSum(z[(r,start,sec,s)] for (r,start,sec2,s) in z if sec2==sec and start<=t<=start+2)
        ) <= 1

# Subject uniqueness
for sub in lecture_subjects:
    for t in timeslots:
        model += lpSum(x[(r,t,sec,sub)]
                       for (r,t2,sec,s) in x if s==sub and t2==t) <= 1

for sub in lab_subjects:
    for t in timeslots:
        model += lpSum(z[(r,start,sec,sub)]
                       for (r,start,sec,s) in z if s==sub and start<=t<=start+2) <= 1

# Lecture units
for sec in sections:
    for sub in lecture_subjects:
        model += lpSum(x[(r,t,sec,sub)]
                       for (r,t,sec2,s) in x if sec2==sec and s==sub) == lecture_subjects[sub]

# Lab units
for sec in sections:
    for sub in lab_subjects:
        model += lpSum(z[(r,t,sec,sub)]
                       for (r,t,sec2,s) in z if sec2==sec and s==sub) == lab_subjects[sub]

# Lunch break (slot 5 every day)
lunch_slots = [5 + 9*i for i in range(5)]

for (r,t,sec,sub) in x:
    if t in lunch_slots:
        model += x[(r,t,sec,sub)] == 0

for (r,t,sec,sub) in z:
    if t in lunch_slots or t+1 in lunch_slots or t+2 in lunch_slots:
        model += z[(r,t,sec,sub)] == 0

# ----------------------------
# OBJECTIVE FUNCTION
# ----------------------------

# Vacant penalty
vacant_penalty = lpSum(
    (1 - y[(sec,t)])
    for sec in sections
    for t in timeslots
)

# Balance penalty
balance_penalty = []
for sec in sections:
    avg = sum(lecture_subjects.values()) / 5
    for d in range(5):
        slots = range(1 + d*9, 10 + d*9)
        load = lpSum(y[(sec,t)] for t in slots)

        p = LpVariable(f"bal_{sec}_{d}", lowBound=0)

        model += p >= load - avg
        model += p >= avg - load

        balance_penalty.append(p)

# Final objective
model += 15 * vacant_penalty + 10 * lpSum(balance_penalty)

# ----------------------------
# SOLVE
# ----------------------------

model.solve()
print("Status:", LpStatus[model.status])

# ----------------------------
# OUTPUT
# ----------------------------

schedule = []

for (r,t,sec,sub) in x:
    if value(x[(r,t,sec,sub)]) == 1:
        schedule.append([sec, sub, r, t])

for (r,t,sec,sub) in z:
    if value(z[(r,t,sec,sub)]) == 1:
        schedule.append([sec, sub+" LAB", r, t])

schedule_df = pd.DataFrame(schedule, columns=["Section","Subject","Room","Time"])

# ----------------------------
# EXCEL TIMETABLE
# ----------------------------

days = ["Mon","Tue","Wed","Thu","Fri"]
times = ["8-9","9-10","10-11","11-12","12-1","1-2","2-3","3-4","4-5"]

def decode(t):
    return days[(t-1)//9], times[(t-1)%9]

for sec in sections:
    grid = {time: {day:"" for day in days} for time in times}

    for _, row in schedule_df.iterrows():
        if row["Section"] == sec:
            d, t = decode(row["Time"])
            grid[t][d] = f"{row['Subject']} ({row['Room']})"

    df_grid = pd.DataFrame(grid).T
    df_grid.to_excel(f"{sec}_timetable.xlsx")

print("Excel timetables generated.")