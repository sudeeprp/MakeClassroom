from openpyxl import load_workbook
import sys
import json
import string

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
        "class_name": string.capwords(classroom_sheet[classroom_col_map['Class name'] + content_row]),
        "school_name": string.capwords(classroom_sheet[classroom_col_map['School name'] + content_row]),
        "locality_10": string.capwords(classroom_sheet[classroom_col_map['Inspection'] + content_row]),
        "locality_20": string.capwords(classroom_sheet[classroom_col_map['Region'] + content_row]),
        "locality_30": string.capwords(classroom_sheet[classroom_col_map['District'] + content_row])
    }
    return classroom_details, class_id

def readTeachers(teachers_sheet):
    teachers = {}
    teacher_row = 2
    teacher_col_map = map_headings(teachers_sheet)
    while teachers_sheet[teacher_col_map['Teacher ID'] + str(teacher_row)] is not None:
        teacher_id = teachers_sheet[teacher_col_map['Teacher ID'] + str(teacher_row)]
        teachers[teacher_id + "/name"] = string.capwords(teachers_sheet[teacher_col_map['Teacher name'] + str(teacher_row)])
        teacher_row += 1
    return teachers

def getDateOfBirth(dob_value):
    dateOfBirth = {"dd": "", "mm": "", "yyyy": dob_value}
    if hasattr(dob_value, 'day'): dateOfBirth["dd"] = str(dob_value.day)
    if hasattr(dob_value, 'month'): dateOfBirth["mm"] = str(dob_value.month)
    if hasattr(dob_value, 'year'): dateOfBirth["yyyy"] = str(dob_value.year)
    return dateOfBirth

def readStudents(students_sheet):
    students = {}
    student_row = 2
    student_col_map = map_headings(students_sheet)
    while students_sheet[student_col_map['Student ID'] + str(student_row)] is not None:
        student_id = students_sheet[student_col_map['Student ID'] + str(student_row)]
        dateOfBirth = getDateOfBirth(students_sheet.wsheet[student_col_map['Date of Birth'] + str(student_row)].value)
        students[student_id + "/first_name"] = string.capwords(students_sheet[student_col_map['First name'] + str(student_row)])
        students[student_id + "/surname"] = string.capwords(students_sheet[student_col_map['Surname'] + str(student_row)])
        students[student_id + "/birth_date/dd"] = dateOfBirth["dd"]
        students[student_id + "/birth_date/mm"] = dateOfBirth["mm"]
        students[student_id + "/birth_date/yyyy"] = dateOfBirth["yyyy"]
        students[student_id + "/gender"] = students_sheet[student_col_map['Gender'] + str(student_row)]
        students[student_id + "/grade"] = str(students_sheet[student_col_map['Grade'] + str(student_row)])
        student_row += 1
    return students

def readClassroomAndAssets(excel_file):
    classroom_and_assets = {}
    w = load_workbook(excel_file)
    classroom_and_assets['classroom_details'], classroom_and_assets['class_id'] = readClassroom(Sheet(w['Classroom']))
    classroom_and_assets['teachers'] = readTeachers(Sheet(w['Teachers']))
    classroom_and_assets['students'] = readStudents(Sheet(w['Students']))

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
