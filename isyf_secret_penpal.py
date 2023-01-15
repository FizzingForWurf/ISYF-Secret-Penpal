from xlutils.copy import copy
import xlwt
import xlrd
import os.path
import sys
import random

NAME_COL = 0
SCHOOL_COL = 1
NATIONALITY_COL = 2
GENDER_COL = 3
GROUP_COL = 4

NO_OF_PPL, NO_OF_GRP = 0, 0 # Update in prompt_user
error_ppl, error_grp = False, False # Update in map_groupings

local_schools = [] # Populate in map_groupings()
original_delegates = [] # Store original sequence of delegates from input file

OUTPUT_FILE = "pairing_results.xls"
results_wb = xlwt.Workbook()

class Student:
    def __init__(self, name, school, nationality, gender, group, isForeigner):
        self.name = name
        self.school = school
        self.nationality = nationality
        self.gender = gender
        self.group = group
        self.isForeigner = isForeigner

def prompt_user():
    print('''
===================== ISYF Secret Penpal delegates pairing =====================
======================== Done by: Tong Zheng Hong 18S6B ========================''')
    print("")
    print('''Welcome! This program is written to allocate pairs given a list of student
information for the Secret Penpal activity. Individuals will be paired up if
they are from different groups and schools. Foreign students will always be
paired with locals as well to promote culture appreciation amongst the
participants.

In order for the program to run smoothly, the input file needs to be in a
certain format which is listed below. The information here only serves as a
GUIDE, for more information and troubleshooting, please visit: 
https://github.com/FizzingForWurf/ISYF-Secret-Penpal

Alternatively, you can contact me at my Telegram handle @zheng_hongg
or email me at zhtong@gmail.com!''')
    input("\nPlease press ENTER to continue...")

    print("")
    print('''The FIRST SHEET in the input file should contain all LOCAL schools.
It should start at the top right cell. Please do NOT include other information
other than local schools in the first column.''')
    input("\nPlease press ENTER to continue...")

    print("")
    print('''The SECOND SHEET should contain the delegates information in their
respective groups.
1. The first ROW contains the headers (Name, School, Nationality, Gender)
This row of information will be ignored by the program
2. First COLUMN: Name of the participant
3. Second COLUMN: School
4. Third COLUMN: Nationality
5. Forth COLUMN: Gender
6. Important: Leave an EMPTY ROW between participants of different groups
(for the program to identify the groupings)''')
    print('''\nExample:
Name		School		Nationality		Gender
Tom		ABC		Singaporean		M
Jane		BCD		Singaporean		F
Mary		DEF		Malaysian		F

Jerry		ABC		Singaporean		M
Jeff		BCD		Malaysian		M

Sarah		DEF		Singaporean		F
Sam		ABC		Singaporean		M

First group: Tom, Jane and Mary
Second group: Jerry and Jeff
Third group: Sarah and Sam''')
    input("\nPlease press ENTER to continue...")

    print("")
    print('''Any other sheets should be placed BEHIND the first two sheets and will be
ignored by the program''')
    input("\nPlease press ENTER to continue...")

    print("")
    print('''Lastly, please ensure that the input file is in the same folder as the
exe program file!''')
    input_file = input("\nEnter the name of the file (exclude .xlsx): ") + ".xlsx"
    while (not os.path.isfile(input_file)):
        print("'" + input_file + "' does NOT exists! Please ensure that the input file is in the same folder as the exe program file!")
        input_file = input("Enter the name of the file (exclude .xlsx): ") + ".xlsx"

    global NO_OF_PPL, NO_OF_GRP
    NO_OF_PPL = int(input("Enter the number of PARTICIPANTS: "))
    NO_OF_GRP = int(input("Enter the number of GROUPS: "))
    
    delegates_wb = xlrd.open_workbook(input_file)
    return delegates_wb

def write_header(sheet, col_offset=0):
    sheet.write(0, NAME_COL + col_offset, "Name")
    sheet.write(0, SCHOOL_COL + col_offset, "School")
    sheet.write(0, NATIONALITY_COL + col_offset, "Nationality")
    sheet.write(0, GENDER_COL + col_offset, "Gender")
    sheet.write(0, GROUP_COL + col_offset, "Group")

def write_student_info(sheet, row, student, col_offset=0):
    sheet.write(row, NAME_COL + col_offset, student.name)
    sheet.write(row, SCHOOL_COL + col_offset, student.school)
    sheet.write(row, NATIONALITY_COL + col_offset, student.nationality)
    sheet.write(row, GENDER_COL + col_offset, student.gender)
    sheet.write(row, GROUP_COL + col_offset, student.group)

def map_groupings(delegates_wb):
    local_school_sheet = delegates_wb.sheet_by_index(0) # First sheet should contain local schools
    grouping_sheet = delegates_wb.sheet_by_index(1) # Second sheet should contain delegate groupings

    # Get local schools in first sheet
    for row in range(local_school_sheet.nrows):
        school = local_school_sheet.cell_value(row, 0) # Get school in first column
        school = school.upper().strip()
        local_schools.append(school)

    group_counter, skipped = 1, False
    school_dict = {} # Stores students per school (school -> list of student objects)
    for row in range(1, grouping_sheet.nrows):
        if (grouping_sheet.cell_value(row, 0) == ''):
            if (not skipped):
                group_counter += 1
                
            skipped = True
            continue

        skipped = False
        school = grouping_sheet.cell_value(row, SCHOOL_COL)
        school = school.upper().strip()

        name = grouping_sheet.cell_value(row, NAME_COL)
        nationality = grouping_sheet.cell_value(row, NATIONALITY_COL)
        gender = grouping_sheet.cell_value(row, GENDER_COL)
        isForeigner = (school not in local_schools)
        person = Student(name, school, nationality, gender, group_counter, isForeigner)
        original_delegates.append(person)

        if (school in school_dict):
            school_dict[school].append(person)
        else:
            school_dict[school] = [person]

    # Check if number of people and group correct
    global error_ppl, error_grp
    error_ppl = NO_OF_PPL != len(original_delegates)
    error_grp = NO_OF_GRP != group_counter

    # Sort schools by the number of students
    # Students from schools with more people should on top
    processed_delegates = sorted(school_dict.items(), key=lambda x: len(x[1]), reverse=True)
    foreign, local = [], []
    for sch, students in processed_delegates:
        if (sch in local_schools):
            local += students
        else:
            foreign += students
    
    new_delegates_sheet = results_wb.add_sheet("Delegates")
    write_header(new_delegates_sheet)

    new_delegates = foreign + local;
    for i in range(1, len(new_delegates)+1):
        student = new_delegates[i-1]
        write_student_info(new_delegates_sheet, i, student)

    results_wb.save(OUTPUT_FILE)
    return new_delegates

def allocate_pairs(delegates):
    print("\nAllocating pairs... \n")
    
    pairs_sheet = results_wb.add_sheet("Pairs")
    write_header(pairs_sheet)
    write_header(pairs_sheet, col_offset=6)

    row_count = 1
    pair_result = {}
    while (len(delegates) > 1):
        pair_index = random.randint(1, len(delegates)-1)
        cur_stu, pair_stu = delegates[0], delegates[pair_index]
        print("%-25s%3s%25s" % (cur_stu.name, "-->", pair_stu.name), end='')

        both_foreigner = cur_stu.isForeigner and pair_stu.isForeigner
        same_school = cur_stu.school == pair_stu.school
        same_group = cur_stu.group == pair_stu.group

        if (not same_school and not same_group and not both_foreigner):
            # Record down details of the pair
            pair_result[cur_stu.name] = pair_stu
            pair_result[pair_stu.name] = cur_stu

            write_student_info(pairs_sheet, row_count, cur_stu)
            write_student_info(pairs_sheet, row_count, pair_stu, col_offset=6)

            #REMOVE first person and the penpal from the lists
            delegates.pop(pair_index)
            delegates.pop(0)
            
            row_count += 1
            print(" (OK!)")
            continue
        
        print(" (NO MATCH)")

    print("\nNumber of pairs:", row_count-1)
    if (len(delegates) > 0): # There is still someone left
        print("The last odd person is: " + delegates[0].name)
        print("Please manually add this person to an existing pair!")
    else:
        print("No odd person left! Everyone is paired up :)")

    print("\nSuccessfully allocated all pairs!")
    print("Pairing results can be found in pairing_results.xls file!")
    results_wb.save(OUTPUT_FILE)
    return pair_result

def show_in_groups(pairs):
    grouped_sheet = results_wb.add_sheet("Grouped")
    write_header(grouped_sheet)
    write_header(grouped_sheet, col_offset=6)

    for i, student in enumerate(original_delegates):
        write_student_info(grouped_sheet, i+1, student)
        
        if student.name in pairs:
            other = pairs[student.name]
            write_student_info(grouped_sheet, i+1, other, col_offset=6)
    
    results_wb.save(OUTPUT_FILE)

delegates_wb = prompt_user()
delegates = map_groupings(delegates_wb)

if (error_ppl):
    print("\nERROR: Incorrect number of participants detected!")
if (error_grp):
    print("\nERROR: Incorrect number of groups detected!")
if (error_ppl or error_grp):
    print('''
Please ensure that the number of participants and/or groups entered
is correct or check the format of the input excel sheet.

For more information, please visit: 
https://github.com/FizzingForWurf/ISYF-Secret-Penpal

Alternatively, contact me at my Telegram handle @zheng_hongg or
email me at zhtong@gmail.com''' )
    
    input("\nPress ENTER to exit the program...")
    sys.exit()

pairs = allocate_pairs(delegates)
show_in_groups(pairs)

input("\nPress ENTER to exit the program...")