import random
from os import path
import pandas as pd
from datetime import date
import openpyxl
from openpyxl.styles import PatternFill

# ------------------------------------------- GLOBALS ------------------------------------------- #

today = date.today().strftime("%d/%m/%y")

waiters = ['Lesly', 'Ismail', 'Alex', 'Hakan', 'Alexandra','Yolanda', 'Pragya', 'Maroon', 'Sivu', 'Kevin', 'Patrick']
sections = {
    'Table1': {'location': 'outside', 'capacity': 1, 'importance': True},
    'Table2-3': {'location': 'outside', 'capacity': 1, 'importance': True},
    'Table4-5': {'location': 'outside', 'capacity': 1, 'importance': True},
    'NewSection': {'location': 'outside', 'capacity': 3, 'importance': True},
    'Middle': {'location': 'inside', 'capacity': 2, 'importance': False},
    '13-14-15': {'location': 'inside', 'capacity': 1, 'importance': True},
    '16-17-18': {'location': 'inside', 'capacity': 1, 'importance': True},
    'SmokingSection': {'location': 'inside', 'capacity': 1, 'importance': False},
}
section_assignments = dict.fromkeys(sections.keys(), None)
unimportant_sections = []
unimportant_sections_dict = {}
important_sections = []
important_sections_dict = {}
outside_sections = []
inside_sections = []
for section_key, section_value in sections.items():
    if section_value["location"] == "outside":
        outside_sections.append(section_key)
    else:
        inside_sections.append(section_key)
    if section_value["importance"]:
        important_sections.append(section_key)
        important_sections[section_key] = section_value
    else:
        unimportant_sections.append(section_key)
        unimportant_sections[section_key] = section_value
        
try:
    waiter_history_df = pd.read_excel("./Data/RotationCopy.xlsx", index_col=0)
    waiter_history_dict = waiter_history_df.to_dict("list")
except:
    if not path.exists("./Data/Rotation.xlsx"):
        rotation = pd.DataFrame(columns=waiters)
        rotation.index.name = "Date"
        rotation.to_excel("./Data/Rotation.xlsx")
print(waiter_history_dict)
# ------------------------------------------- FUNCTIONS ------------------------------------------- #

def CheckAvailableWaiter(section_key, section_value):
    desired_section_count = {}
    shuffled_waiter_history_dict = {k: waiter_history_dict[k] for k in random.sample(sorted(waiter_history_dict.keys()), len(waiter_history_dict))}
    for waiter, history in shuffled_waiter_history_dict.items():
        if (sections[history[-1]]["location"] == section_value["location"]) or not waiter in waiters:
            continue
        history_count = history.count(section_key)
        desired_section_count.update({f"{waiter}" : history_count})
    sorted_desired_section_count = sorted(desired_section_count.items(), key=lambda x:x[1], reverse=True)
        
    try:
        return sorted_desired_section_count[-1][0]
    except:
        pass

def RotationLog(section_assignments):
    waiter_assignments = {}
    for section_key, section_value in section_assignments.items():
        if section_value == None:
            section_value = []
        for value in section_value:
            waiter_assignments.update({f"{value}": section_key})
    new_row = pd.DataFrame(waiter_assignments, index=[today])
    rotation_df = pd.read_excel("./Data/RotationCopy.xlsx", index_col=0)
    rotation_df.index.name = "Date"
    rotation_df = pd.concat([rotation_df, new_row])
    rotation_df.to_excel("./Data/RotationCopy.xlsx", index_label="Date", sheet_name="Cycle")
    styleWorkbook = openpyxl.load_workbook("./Data/RotationCopy.xlsx")
    styleSheet = styleWorkbook["Cycle"]
    for column in styleSheet.iter_cols():
        for cell in column:
            if cell.value not in unimportant_sections:
                if cell.value in inside_sections:
                    styleSheet[cell.coordinate].fill = PatternFill(start_color="AEF359", end_color="AEF359", fill_type="solid")
                elif cell.value in outside_sections:
                    styleSheet[cell.coordinate].fill = PatternFill(start_color="BC544B", end_color="BC544B", fill_type="solid")
    styleWorkbook.save("./Data/RotationCopy.xlsx")

# ------------------------------------------- MAIN ------------------------------------------- #

while waiters:
    for section_key, section_value in important_sections_dict.items():
        if section_value["capacity"] == 0:
            continue
        if not waiters:
            break
        waiter = CheckAvailableWaiter(section_key, section_value)
        if section_value["capacity"] > 0 and section_assignments[f"{section_key}"] == None:
            section_assignments.update({f"{section_key}": [waiter]})
            waiters.remove(waiter)
            section_value["capacity"] -= 1
            continue
        elif section_value["capacity"] > 0 and section_assignments[f"{section_key}"] != None:
            already_assigned = section_assignments[f"{section_key}"]
            already_assigned.append(waiter)
            section_assignments.update({f"{section_key}": already_assigned})
            waiters.remove(waiter)
            section_value['capacity'] -= 1
            continue
        else:
            break
    for section_key, section_value in unimportant_sections_dict.items():
        if section_value["capacity"] == 0:
            continue
        if not waiters:
            break
        waiter = CheckAvailableWaiter(section_key, section_value)
        if section_value["capacity"] > 0 and section_assignments[f"{section_key}"] == None:
            section_assignments.update({f"{section_key}": [waiter]})
            waiters.remove(waiter)
            section_value["capacity"] -= 1
            continue
        elif section_value["capacity"] > 0 and section_assignments[f"{section_key}"] != None:
            already_assigned = section_assignments[f"{section_key}"]
            already_assigned.append(waiter)
            section_assignments.update({f"{section_key}": already_assigned})
            waiters.remove(waiter)
            section_value['capacity'] -= 1
            continue
        else:
            break


RotationLog(section_assignments)