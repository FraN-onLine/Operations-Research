import pandas as pd

rooms = ["R101", "R102", "R103"]

timeslots = list(range(1, 46))  # 9 slots per day

sections = ["CS3A", "CS3B"]

subjects = {
    "Linear Algebra": 3,
    "Software Engineering": 3,
    "Operations Research": 3,
    "Automata Theory": 3
}

# who can teach what
instructors = {
    "Reynera": ["Linear Algebra"],
    "Dina": ["Linear Algebra"],
    "Nina": ["Operations Research"],
    "Noel": ["Software Engineering"],
    "Rey":["Software Engineering", "Automata Theory"]
}

section_subjects = {
    "CS3A": ["Linear Algebra", "Software Engineering", "Operations Research", "Automata Theory"],
    "CS3B": ["Linear Algebra", "Software Engineering", "Operations Research", "Automata Theory"]
}

import pandas as pd

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
                            "units": subjects[subject],
                            "go_or_no_go": 0
                        })

df = pd.DataFrame(rows)
print(df)

from pulp import *

model = LpProblem("Class_Scheduling", LpMinimize)

x = {}

for idx, row in df.iterrows():
    x[idx] = LpVariable(f"x_{idx}", cat="Binary")