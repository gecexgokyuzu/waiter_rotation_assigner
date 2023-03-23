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
    'SmokingSection': {'location': 'inside', 'capacity': 2, 'importance': False},
}
unimportant_sections = []
important_sections = []
outside_sections = []
inside_sections = []
full_sections = []
important_section_capacity = 0
unimportant_section_capacity = 0
section_assignments = dict.fromkeys(sections.keys(), None)
assigned_waiters = []
for sectionKey, sectionValues in sections.items():
    if sectionValues["location"] == "outside":
        outside_sections.append(sectionKey)
    else:
        inside_sections.append(sectionKey)
for sectionKey, sectionValues in sections.items():
    if not sectionValues["importance"]:
        unimportant_sections.append(sectionKey)
        unimportant_section_capacity += sectionValues['capacity']
    else:
        important_section_capacity += sectionValues["capacity"]
        important_sections.append(sectionKey)
try:
    waiter_historyDF = pd.read_excel("./Data/RotationCopy.xlsx", index_col=0)
    waiter_historyDict = waiter_historyDF.to_dict("list")
except:
# if rotation.xlsx doesn't exist, create one.
    if not path.exists("./Data/Rotation.xlsx"):
        rotation = pd.DataFrame(columns=waiters)
        rotation.index.name = "Date"
        rotation.to_excel("./Data/Rotation.xlsx")

i = 0

# ------------------------------------------- FUNCTIONS ------------------------------------------- #

def RotationLog(section_assignments):
    waiter_assignments = {}
    for sectionKey, sectionValues in section_assignments.items():
        if sectionValues == None:
            sectionValues = []
        for value in sectionValues:
            waiter_assignments.update({f"{value}": sectionKey})
    new_row = pd.DataFrame(waiter_assignments, index=[today])  # todo: pd.Timestamp(today)
    rotationDF = pd.read_excel("./Data/RotationCopy.xlsx", index_col=0)
    rotationDF.index.name = "Date"
    rotationDF = pd.concat([rotationDF, new_row])
    rotationDF.to_excel("./Data/RotationCopy.xlsx", index_label="Date", sheet_name="Cycle")
    
    # -------------------- color the cells -------------------- #

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

def CheckAvailability(waiter, sectionValues, wantedSec, unimportant_sections):
    lastSecLoc = sections[waiter_historyDict[waiter][-1]]["location"]
    wantedSecLoc = sectionValues["location"]
    waiter_historyList = waiter_historyDict[waiter]
    waiter_historyList.reverse()
    for section in unimportant_sections:
        if section in waiter_historyList and section not in waiter_historyList[-3:]:
            waiter_historyList.remove(section)
    if lastSecLoc != wantedSecLoc and wantedSec not in waiter_historyList:
        return True
    else:
        return False
    
# ------------------------------------------- MAIN ------------------------------------------- #

while len(assigned_waiters) < (important_section_capacity + unimportant_section_capacity) and i < 40:
    random.shuffle(waiters)
    i += 1
    for section_key, section_values in sections.items():
        for waiter in waiters:
            if waiter in assigned_waiters:
                continue
            elif section_key in full_sections:
                continue
            if not CheckAvailability(waiter, section_values, section_key, unimportant_sections):
                continue
            if sectionValues["capacity"] > 0 and section_assignments[f"{sectionKey}"] == None:
                section_assignments.update({f"{sectionKey}": [waiter]})
                assigned_waiters.append(waiter)
                sectionValues["capacity"] -= 1
                continue
            elif sectionValues["capacity"] > 0 and section_assignments[f"{sectionKey}"] != None:
                already_assigned = section_assignments[f"{sectionKey}"]
                already_assigned.append(waiter)
                section_assignments.update({f"{sectionKey}": already_assigned})
                assigned_waiters.append(waiter)
                sectionValues['capacity'] -= 1
                continue
            else:
                break