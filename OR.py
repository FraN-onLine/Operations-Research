import pandas as pd
from pulp import *
from collections import defaultdict

# ----------------------------
# DATA
# ----------------------------

lecture_rooms = ["R100A", "R100B", "R100C"]
lab_rooms = ["Lab1", "Lab2"]

timeslots = list(range(1, 46))  # 5 days x 9
sections = ["CS1A","CS3A", "CS3B"]

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
    "Database Systems": 3,
    "Operating System": 3,
    "Data Structures & Algorithms": 3,
    "Introduction to Game Development": 3,
    "Ethics": 3,
    "Fundamentals of Accounting for IT": 3,
    "Environmental Science": 3,
    "Dance": 2,
}

lab_subjects = {
    "Computer Programming I Lab": 1,
    "Computer Science Fundamentals Lab": 1,
    "Object Oriented Programming Lab": 1,
    "Computer Programming II Lab": 1,
    "Data Structures Lab": 1,
    "Digital Design Lab": 1,
    "Database Systems Lab": 1,
    "Discre"
    "Distributed Systems Lab": 1,
    "Software Engineering Lab": 1,
    "Information Assurance Lab": 1,
    
    #IT subjects
    "Introduction to Computing Lab": 1, 
    "Computer Programming 1 Lab": 1,
    "Operating System Lab": 1,
}

#FOR FIRST SEMESTER of 2026-2027
section_subjects = {
    "CS1A": ["Understanding the Self"],
    "CS3A": ["Linear Algebra", "Software Engineering", "Operations Research", "Automata Theory", "Distributed Systems", "Information Assurance", "Information Assurance Lab", "Software Engineering Lab", "Distributed Systems Lab"],
    "CS3B": ["Linear Algebra", "Software Engineering", "Operations Research", "Automata Theory", "Distributed Systems", "Information Assurance", "Information Assurance Lab", "Software Engineering Lab", "Distributed Systems Lab"]
}

# Precompute valid lab starts (same-day guarantee)
valid_lab_starts = [1,4,7,10,13,16,19,22,25,28,31,34,37,40,43]

# ----------------------------
# MODEL
# ----------------------------

model = LpProblem("Scheduling", LpMinimize)

# ----------------------------
# VARIABLES
# ----------------------------

x = LpVariable.dicts("x",
    [(r,t,sec,sub)
     for r in lecture_rooms
     for t in timeslots
     for sec in sections
     for sub in lecture_subjects
     if sub in section_subjects[sec]],
    cat="Binary"
)

z = LpVariable.dicts("z",
    [(r,t,sec,sub)
     for r in lab_rooms
     for t in valid_lab_starts
     for sec in sections
     for sub in lab_subjects
     if sub in section_subjects[sec]],
    cat="Binary"
)

y = LpVariable.dicts("y",
    [(sec,t) for sec in sections for t in timeslots],
    cat="Binary"
)

# ----------------------------
# INDEXING (SPEED BOOST)
# ----------------------------

x_by_room_time = defaultdict(list)
x_by_section_time = defaultdict(list)
x_by_subject_time = defaultdict(list)
x_by_section_subject = defaultdict(list)

for (r,t,sec,sub) in x:
    x_by_room_time[(r,t)].append(x[(r,t,sec,sub)])
    x_by_section_time[(sec,t)].append(x[(r,t,sec,sub)])
    x_by_subject_time[(sub,t)].append(x[(r,t,sec,sub)])
    x_by_section_subject[(sec,sub)].append(x[(r,t,sec,sub)])

z_by_room_time = defaultdict(list)
z_by_section_time = defaultdict(list)
z_by_subject_time = defaultdict(list)
z_by_section_subject = defaultdict(list)

for (r,t,sec,sub) in z:
    for dt in range(3):
        z_by_room_time[(r,t+dt)].append(z[(r,t,sec,sub)])
        z_by_section_time[(sec,t+dt)].append(z[(r,t,sec,sub)])
        z_by_subject_time[(sub,t+dt)].append(z[(r,t,sec,sub)])
    z_by_section_subject[(sec,sub)].append(z[(r,t,sec,sub)])

# ----------------------------
# CONSTRAINTS
# ----------------------------

# Link y (occupied slots)
for sec in sections:
    for t in timeslots:
        model += y[(sec,t)] == (
            lpSum(x_by_section_time[(sec,t)]) +
            lpSum(z_by_section_time[(sec,t)])
        )

# Room constraints
for r in lecture_rooms:
    for t in timeslots:
        model += lpSum(x_by_room_time[(r,t)]) <= 1

for r in lab_rooms:
    for t in timeslots:
        model += lpSum(z_by_room_time[(r,t)]) <= 1

# Section constraints
for sec in sections:
    for t in timeslots:
        model += (
            lpSum(x_by_section_time[(sec,t)]) +
            lpSum(z_by_section_time[(sec,t)])
        ) <= 1

# Subject uniqueness
for sub in lecture_subjects:
    for t in timeslots:
        model += lpSum(x_by_subject_time[(sub,t)]) <= 1

for sub in lab_subjects:
    for t in timeslots:
        model += lpSum(z_by_subject_time[(sub,t)]) <= 1

# Units (only for subjects in section_subjects)
for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            model += lpSum(x_by_section_subject[(sec,sub)]) == lecture_subjects[sub]

    for sub in lab_subjects:
        if sub in section_subjects[sec]:
            model += lpSum(z_by_section_subject[(sec,sub)]) == lab_subjects[sub]

# Lunch slots definition
lunch_slots = [5, 14, 23, 32, 41]

# Constraint 0a: NO lecture subject starts at lunch hour
for sub in lecture_subjects:
    for t in lunch_slots:
        model += lpSum(x[(r,t,sec,sub)] for r in lecture_rooms for sec in sections if sub in section_subjects[sec]) == 0

# Constraint 0b: Lunch hours can only be occupied by lab hours (no lectures)
for t in lunch_slots:
    for sec in sections:
        model += lpSum(x[(r,t,sec,sub)] for r in lecture_rooms for sub in lecture_subjects if sub in section_subjects[sec]) == 0

# ----------------------------
# NEW CONSTRAINTS
# ----------------------------

# Constraint 1: Same room for all sessions of a subject in a section (Lectures)
for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            for r1 in lecture_rooms:
                for r2 in lecture_rooms:
                    if r1 < r2:
                        # All sessions of this subject must be in the same room
                        model += (lpSum(x[(r1,t,sec,sub)] for t in timeslots) + 
                                 lpSum(x[(r2,t,sec,sub)] for t in timeslots)) <= lecture_subjects[sub]

# Constraint 2: Same room for all sessions of a subject in a section (Labs)
for sec in sections:
    for sub in lab_subjects:
        if sub in section_subjects[sec]:
            for r1 in lab_rooms:
                for r2 in lab_rooms:
                    if r1 < r2:
                        # All sessions of this subject must be in the same room
                        model += (lpSum(z[(r1,t,sec,sub)] for t in valid_lab_starts) + 
                                 lpSum(z[(r2,t,sec,sub)] for t in valid_lab_starts)) <= lab_subjects[sub]

# Constraint 3: Once per day unless consecutive (Lectures)

for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            for day in range(5):
                day_slots = [t for t in range(1 + day*9, 10 + day*9) if t not in lunch_slots]
                # For non-consecutive slots, at most one can be used
                for i in range(len(day_slots)):
                    for j in range(i+2, len(day_slots)):
                        t1, t2 = day_slots[i], day_slots[j]
                        model += (lpSum(x[(r,t1,sec,sub)] for r in lecture_rooms) + 
                                 lpSum(x[(r,t2,sec,sub)] for r in lecture_rooms)) <= 1

# Constraint 4: Once per day unless consecutive (Labs)
for sec in sections:
    for sub in lab_subjects:
        if sub in section_subjects[sec]:
            for day in range(5):
                day_slots = [t for t in range(1 + day*9, 10 + day*9) if t not in lunch_slots]
                # For labs, check the starting slots (3-hour blocks)
                lab_day_starts = [t for t in day_slots if t in valid_lab_starts]
                for i in range(len(lab_day_starts)):
                    for j in range(i+1, len(lab_day_starts)):
                        t1, t2 = lab_day_starts[i], lab_day_starts[j]
                        # If both are scheduled, they must be consecutive (t2 = t1 + 3)
                        if t2 != t1 + 3:
                            model += (lpSum(z[(r,t1,sec,sub)] for r in lab_rooms) + 
                                     lpSum(z[(r,t2,sec,sub)] for r in lab_rooms)) <= 1

# Lunch break (slot 5 each day)

for t in lunch_slots:
    for sec in sections:
        model += y[(sec,t)] == 0

# ----------------------------
# OBJECTIVE
# ----------------------------

# Vacant minimization
vacant_penalty = lpSum(1 - y[(sec,t)] for sec in sections for t in timeslots)

# Balance penalty
balance_penalty = []
avg = sum(lecture_subjects.values()) / 5

for sec in sections:
    for d in range(5):
        slots = range(1 + d*9, 10 + d*9)
        load = lpSum(y[(sec,t)] for t in slots)

        p = LpVariable(f"bal_{sec}_{d}", lowBound=0)
        model += p >= load - avg
        model += p >= avg - load
        balance_penalty.append(p)

# Preference: Same time slot across days (Lectures)
same_timeslot_penalty_lecture = []
for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            # Get timeslot within each day (0-8, excluding lunch)
            for timeslot_in_day in range(1, 10):  # 1-9 within a day
                if timeslot_in_day != 5:  # Skip lunch slot
                    days_using_this_slot = []
                    for day in range(5):
                        t = 1 + day*9 + timeslot_in_day - 1
                        uses_slot = lpSum(x[(r,t,sec,sub)] for r in lecture_rooms)
                        days_using_this_slot.append(uses_slot)
                    # Penalty: count mismatches (penalize if not all days use same slot or same pattern)
                    for i in range(len(days_using_this_slot)-1):
                        p = LpVariable(f"same_time_lec_{sec}_{sub}_{timeslot_in_day}_{i}", lowBound=0)
                        model += p >= days_using_this_slot[i] - days_using_this_slot[i+1]
                        model += p >= days_using_this_slot[i+1] - days_using_this_slot[i]
                        same_timeslot_penalty_lecture.append(p)

# Preference: Same time slot across days (Labs)
same_timeslot_penalty_lab = []
for sec in sections:
    for sub in lab_subjects:
        if sub in section_subjects[sec]:
            # Get starting timeslot within each day from valid_lab_starts
            lab_slots_per_day = {}
            for day in range(5):
                lab_slots_per_day[day] = [t for t in valid_lab_starts if 1 + day*9 <= t < 10 + day*9]
            
            if lab_slots_per_day[0]:  # If there are valid lab slots
                for slot_idx, timeslot in enumerate(lab_slots_per_day[0]):
                    days_using_this_slot = []
                    for day in range(5):
                        if slot_idx < len(lab_slots_per_day[day]):
                            t = lab_slots_per_day[day][slot_idx]
                            uses_slot = lpSum(z[(r,t,sec,sub)] for r in lab_rooms)
                            days_using_this_slot.append(uses_slot)
                    if len(days_using_this_slot) > 1:
                        for i in range(len(days_using_this_slot)-1):
                            p = LpVariable(f"same_time_lab_{sec}_{sub}_{slot_idx}_{i}", lowBound=0)
                            model += p >= days_using_this_slot[i] - days_using_this_slot[i+1]
                            model += p >= days_using_this_slot[i+1] - days_using_this_slot[i]
                            same_timeslot_penalty_lab.append(p)

# Preference: Same room for subject (Lectures)
same_room_penalty_lecture = []
for sec in sections:
    for sub in lecture_subjects:
        if sub in section_subjects[sec]:
            for r1 in lecture_rooms:
                for r2 in lecture_rooms:
                    if r1 < r2:
                        uses_r1 = lpSum(x[(r1,t,sec,sub)] for t in timeslots)
                        uses_r2 = lpSum(x[(r2,t,sec,sub)] for t in timeslots)
                        # Penalty if both rooms are used
                        p = LpVariable(f"same_room_lec_{sec}_{sub}_{r1}_{r2}", lowBound=0)
                        model += p >= uses_r1 + uses_r2 - 1
                        same_room_penalty_lecture.append(p)

# Preference: Same room for subject (Labs)
same_room_penalty_lab = []
for sec in sections:
    for sub in lab_subjects:
        if sub in section_subjects[sec]:
            for r1 in lab_rooms:
                for r2 in lab_rooms:
                    if r1 < r2:
                        uses_r1 = lpSum(z[(r1,t,sec,sub)] for t in valid_lab_starts)
                        uses_r2 = lpSum(z[(r2,t,sec,sub)] for t in valid_lab_starts)
                        # Penalty if both rooms are used
                        p = LpVariable(f"same_room_lab_{sec}_{sub}_{r1}_{r2}", lowBound=0)
                        model += p >= uses_r1 + uses_r2 - 1
                        same_room_penalty_lab.append(p)

model += (15 * vacant_penalty + 
          10 * lpSum(balance_penalty) + 
          5 * lpSum(same_timeslot_penalty_lecture) + 
          5 * lpSum(same_timeslot_penalty_lab) +
          8 * lpSum(same_room_penalty_lecture) +
          8 * lpSum(same_room_penalty_lab))

# ----------------------------
# SOLVE
# ----------------------------

print("Solving...")
model.solve()
print("Status:", LpStatus[model.status])

# ----------------------------
# OUTPUT (EXPAND LABS)
# ----------------------------

schedule = []

for (r,t,sec,sub) in x:
    if value(x[(r,t,sec,sub)]) == 1:
        schedule.append([sec, sub, r, t, "Lecture"])

for (r,t,sec,sub) in z:
    if value(z[(r,t,sec,sub)]) == 1:
        for dt in range(3):
            schedule.append([sec, sub, r, t+dt, "Lab"])

schedule_df = pd.DataFrame(schedule, columns=["Section","Subject","Room","Time","Type"])

# ----------------------------
# EXCEL OUTPUT (COLORED)
# ----------------------------

days = ["Mon","Tue","Wed","Thu","Fri"]
times = ["8-9","9-10","10-11","11-12","12-1","1-2","2-3","3-4","4-5"]

def decode(t):
    return days[(t-1)//9], times[(t-1)%9]

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

        from openpyxl.styles import PatternFill

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

print("Saved: Full_Timetable.xlsx")