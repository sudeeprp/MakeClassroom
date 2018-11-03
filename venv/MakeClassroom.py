from openpyxl import load_workbook
import sys
import json
import string
import random


IDCharacters = "123456789ABCDEFGHJKLMNOPQRSTUVWXYZ"
def unique_id():
    id = ""
    for x in range(8):
        i = random.randint(0, len(IDCharacters)-1)
        id += IDCharacters[i]
    return id

def fill_name(name):
    filled_name = ""
    if name is not None:
        filled_name = name
    return filled_name

def capped_name(name):
    return string.capwords(fill_name(name))

def split_name(students_sheet, student_col_map, student_row):
    surname = capped_name(students_sheet[student_col_map['Surname'] + str(student_row)])
    first_name = capped_name(students_sheet[student_col_map['First name'] + str(student_row)])
    if first_name == "":
        name_parts = surname.split()
        surname = name_parts[0]
        if len(name_parts) > 1:
            first_name = ' '.join(name_parts[1:])
    return first_name, surname

def map_headings(ws, heading_row=1, start_col='A'):
    excel_col_map = {}
    cur_column = start_col
    head_row = str(heading_row)
    col_ord = ord(cur_column)
    while ws[cur_column + head_row] is not None and ws[cur_column + head_row] != "" and cur_column != 'Z':
        excel_col_map[ws[cur_column + head_row]] = chr(col_ord)
        col_ord += 1
        cur_column = chr(col_ord)
    if cur_column == 'Z':
        print('ERROR: more columns than expected!\n')
        return None
    return excel_col_map


def getSingleValue(value):
    singleValue = value
    if type(value) is list or type(value) is tuple:
        singleValue = value[0]
    return singleValue

class Sheet:
    def __init__(self, wsheet):
        self.wsheet = wsheet

    def __getitem__(self, item):
        cell_value = getSingleValue(self.wsheet[item].value)
        if cell_value is not None:
            cell_value = str(self.wsheet[item].value).strip()
        return cell_value


def readClassroom(classroom_sheet):
    content_row = '2'
    classroom_col_map = map_headings(classroom_sheet)
    class_id = classroom_sheet[classroom_col_map['Class ID'] + content_row]
    classroom_details = {
        "class_name": fill_name(classroom_sheet[classroom_col_map['Class name'] + content_row]),
        "school_name": fill_name(classroom_sheet[classroom_col_map['School name'] + content_row]),
        "locality_10": fill_name(classroom_sheet[classroom_col_map['Inspection'] + content_row]),
        "locality_20": fill_name(classroom_sheet[classroom_col_map['Region'] + content_row]),
        "locality_30": fill_name(classroom_sheet[classroom_col_map['District'] + content_row])
    }
    return classroom_details, class_id

def readTeachers(teachers_sheet):
    teachers = {}
    teacher_row = 2
    teacher_col_map = map_headings(teachers_sheet)
    while teachers_sheet[teacher_col_map['Teacher name'] + str(teacher_row)] is not None:
        teacher_id = getId(teachers_sheet, teacher_col_map['Teacher ID'], str(teacher_row))
        teachers[teacher_id + "/name"] = capped_name(teachers_sheet[teacher_col_map['Teacher name'] + str(teacher_row)])
        teacher_row += 1
    return teachers

def getId(sheet, column, student_row):
    id = sheet[column + str(student_row)]
    if id is None or id == "":
        id = unique_id()
    return id

def getDateOfBirth(dob_value):
    if dob_value is None: dob_value = ''
    dateOfBirth = {"dd": "", "mm": "", "yyyy": dob_value}
    if hasattr(dob_value, 'day'): dateOfBirth["dd"] = str(dob_value.day)
    if hasattr(dob_value, 'month'): dateOfBirth["mm"] = str(dob_value.month)
    if hasattr(dob_value, 'year'): dateOfBirth["yyyy"] = str(dob_value.year)
    return dateOfBirth

def readStudents(students_sheet):
    students = {}
    student_row = 2
    student_col_map = map_headings(students_sheet)
    while students_sheet[student_col_map['Surname'] + str(student_row)] is not None:
        student_id = getId(students_sheet, student_col_map['Student ID'], student_row)
        dateOfBirth = getDateOfBirth(students_sheet.wsheet[student_col_map['Date of Birth'] + str(student_row)].value)
        students[student_id + "/first_name"], students[student_id + "/surname"] = \
            split_name(students_sheet, student_col_map, student_row)
        students[student_id + "/birth_date/dd"] = dateOfBirth["dd"]
        students[student_id + "/birth_date/mm"] = dateOfBirth["mm"]
        students[student_id + "/birth_date/yyyy"] = dateOfBirth["yyyy"]
        students[student_id + "/gender"] = students_sheet[student_col_map['Gender'] + str(student_row)]
        students[student_id + "/grade"] = str(students_sheet[student_col_map['Grade'] + str(student_row)])
        student_row += 1
    return students


guestGrades = ["1", "2"]
guestGenders = ["boy", "girl"]
iiu = "GI"

def guestStudents():
    guests = {}
    guestStructure = {
        "first_name": "Entraînement",
        "surname": " ",
        "qualifier": "guest"
    }
    guestIndex = 1
    for grade in guestGrades:
        for gender in guestGenders:
            guestId = "GI" + str(guestIndex)
            for key in guestStructure:
                guests[guestId + "/" + key] = guestStructure[key]
            guests[guestId + "/gender"] = gender
            guests[guestId + "/grade"] = grade
            guestIndex += 1
    return guests

b64guestThumbnail = "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"
def guestThumbnails():
    thumbnails = {}
    guestIndex = 1
    for grade in guestGrades:
        for gender in guestGenders:
            guestId = "GI" + str(guestIndex)
            thumbnails["students/" + guestId + "/thumbnail"] = b64guestThumbnail
            guestIndex += 1
    return thumbnails

def readClassroomAndAssets(excel_file):
    classroom_and_assets = {}
    w = load_workbook(excel_file)
    classroom_and_assets['classroom_details'], classroom_and_assets['class_id'] = readClassroom(Sheet(w['Classroom']))
    classroom_and_assets['teachers'] = readTeachers(Sheet(w['Teachers']))
    classroom_and_assets['students'] = readStudents(Sheet(w['Students']))
    classroom_and_assets['students'].update(guestStudents())
    classroom_and_assets['thumbnails'] = guestThumbnails()
    return classroom_and_assets

def writeJSON(outputfilename, classroom_and_assets):
    outputfile = open(outputfilename, 'w')
    outputfile.write(json.dumps(classroom_and_assets, indent=2))
    outputfile.close()

if len(sys.argv) == 2:
    classroom_and_assets = readClassroomAndAssets(sys.argv[1])
    outputfilename = sys.argv[1][0:sys.argv[1].rfind('.')] + '.json'
    writeJSON(outputfilename, classroom_and_assets)
else:
    print('Make a classroom\n\nUsage: ' + sys.argv[0] + ' <excel filename>')
    print('\nOutput file will have same name as excel with .json extension\n')
