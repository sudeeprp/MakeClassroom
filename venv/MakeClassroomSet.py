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
        filled_name = name.strip()
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


def readClassroomsGuestsAndThumbnails(classroom_sheet):
    content_row = 2
    classroom_col_map = map_headings(classroom_sheet)
    classrooms = {}
    guests = {}
    guestThumbs = {}

    while classroom_sheet[classroom_col_map['UDISE Code'] + str(content_row)] is not None:
        class_id = classroom_sheet[classroom_col_map['UDISE Code'] + str(content_row)]
        class_locprefix = class_id + "/"
        classroom_attributes = {
            class_locprefix + "class_name": "Primary",
            class_locprefix + "school_name": fill_name(classroom_sheet[classroom_col_map['School Name'] + str(content_row)]),
            class_locprefix + "locality_10": fill_name(classroom_sheet[classroom_col_map['Taluka'] + str(content_row)]),
            class_locprefix + "locality_20": fill_name(classroom_sheet[classroom_col_map['District'] + str(content_row)]),
            class_locprefix + "locality_30": fill_name(classroom_sheet[classroom_col_map['Zone'] + str(content_row)])
        }
        classrooms.update(classroom_attributes)
        guests.update(guestStudents(class_locprefix))
        guestThumbs.update(guestStudentThumbs(class_locprefix))
        content_row += 1
    return classrooms, guests, guestThumbs


def readTeachers(teachers_sheet):
    teacher_row = 2
    teacher_col_map = map_headings(teachers_sheet)
    teachers = {}

    while teachers_sheet[teacher_col_map['Teacher name'] + str(teacher_row)] is not None:
        class_id = teachers_sheet[teacher_col_map['UDISE Code'] + str(teacher_row)]
        teacher_id = getId(teachers_sheet, teacher_col_map['Teacher ID'], str(teacher_row))
        teacher_locprefix = class_id + "/teachers/" + teacher_id
        teachers[teacher_locprefix + "/name"] = capped_name(teachers_sheet[teacher_col_map['Teacher name'] + str(teacher_row)])
        teacher_row += 1
    return teachers

def readStudents(students_sheet):
    student_row = 2
    student_col_map = map_headings(students_sheet)
    students = {}

    while students_sheet[student_col_map['Surname'] + str(student_row)] is not None or \
          students_sheet[student_col_map['First name'] + str(student_row)] is not None:
        class_id = students_sheet[student_col_map['UDISE Code'] + str(student_row)]
        student_id = getId(students_sheet, student_col_map['Student ID'], student_row)
        student_locprefix = class_id + "/students/" + student_id
        dateOfBirth = getDateOfBirth(students_sheet.wsheet[student_col_map['Date of Birth'] + str(student_row)].value)
        students[student_locprefix + "/first_name"], students[student_locprefix + "/surname"] = \
            split_name(students_sheet, student_col_map, student_row)
        students[student_locprefix + "/birth_date/dd"] = dateOfBirth["dd"]
        students[student_locprefix + "/birth_date/mm"] = dateOfBirth["mm"]
        students[student_locprefix + "/birth_date/yyyy"] = dateOfBirth["yyyy"]
        students[student_locprefix + "/gender"] = \
            students_sheet[student_col_map['Gender'] + str(student_row)].strip().lower()
        students[student_locprefix + "/grade"] = str(students_sheet[student_col_map['Grade'] + str(student_row)])
        student_row += 1
    return students


guestGrades = ["1", "2"]
guestGenders = ["boy", "girl"]

def guestStudents(class_locprefix):
    guests = {}
    guestStructure = {
        "first_name": "Guest",
        "surname": " ",
        "qualifier": "guest"
    }
    key_prefix = class_locprefix + "students/"
    guestIndex = 1
    for grade in guestGrades:
        for gender in guestGenders:
            guestId = "GI" + str(guestIndex)
            for key in guestStructure:
                guests[key_prefix + "/" + guestId + "/" + key] = guestStructure[key]
            guests[key_prefix + "/" + guestId + "/gender"] = gender
            guests[key_prefix + "/" + guestId + "/grade"] = grade
            guestIndex += 1
    return guests

b64guestThumbnail = "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"
def guestStudentThumbs(class_locprefix):
    thumbnails = {}
    guestIndex = 1
    key_prefix = "classrooms/" + class_locprefix + "students/"
    for grade in guestGrades:
        for gender in guestGenders:
            guestId = "GI" + str(guestIndex)
            thumbnails[key_prefix + guestId + "/thumbnail"] = b64guestThumbnail
            guestIndex += 1
    return thumbnails

def readClassroomAndAssets(excel_file):
    classroom_and_assets = {}
    w = load_workbook(excel_file)
    guestStudents = {}
    if 'Classroom' in w:
        classroom_and_assets['classrooms'], guestStudents, classroom_and_assets['thumbnails'] = \
            readClassroomsGuestsAndThumbnails(Sheet(w['Classroom']))
    if 'Teachers' in w:
        classroom_and_assets['teachers'] = readTeachers(Sheet(w['Teachers']))
    if 'Students' in w:
        classroom_and_assets['students'] = readStudents(Sheet(w['Students']))
    if 'students' in classroom_and_assets:
        classroom_and_assets['students'].update(guestStudents)
    return classroom_and_assets

def writeJSON(outputfilename, classroom_and_assets):
    outputfile = open(outputfilename, 'w')
    outputfile.write(json.dumps(classroom_and_assets, indent=2))
    outputfile.close()

if len(sys.argv) == 2:
    classroom_and_assets = readClassroomAndAssets(sys.argv[1])
    classroom_and_assets['type'] = "consolidated a"
    outputfilename = sys.argv[1][0:sys.argv[1].rfind('.')] + '.json'
    writeJSON(outputfilename, classroom_and_assets)
else:
    print('Make a classroom\n\nUsage: ' + sys.argv[0] + ' <excel filename>')
    print('\nOutput file will have same name as excel with .json extension\n')
