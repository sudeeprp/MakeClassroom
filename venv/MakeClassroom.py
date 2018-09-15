from openpyxl import load_workbook
from openpyxl.utils import datetime
import sys
import json

def map_headings(ws, heading_row=1, start_col='A'):
    excel_col_map = {}
    cur_column = start_col
    head_row = str(heading_row)
    col_ord = ord(cur_column)
    while ws[cur_column + head_row].value is not None and ws[cur_column + head_row].value != "" and cur_column != 'Z':
        excel_col_map[ws[cur_column + head_row].value] = chr(col_ord)
        col_ord += 1
        cur_column = chr(col_ord)
    if cur_column == 'Z':
        print('ERROR: more columns than expected!\n')
        return None
    return excel_col_map


def readClassroom(classroom_worksheet):
    classroom = {}
    content_row = '2'
    classroom_col_map = map_headings(classroom_worksheet)
    class_id = classroom_worksheet[classroom_col_map['Class ID'] + content_row].value
    classroom[class_id] = {
        "class_name": classroom_worksheet[classroom_col_map['Class name'] + content_row].value,
        "school_name": classroom_worksheet[classroom_col_map['School name'] + content_row].value,
        "locality_10": classroom_worksheet[classroom_col_map['Inspection'] + content_row].value,
        "locality_20": classroom_worksheet[classroom_col_map['Region'] + content_row].value,
        "locality_30": classroom_worksheet[classroom_col_map['District'] + content_row].value
    }
    return classroom, class_id

def readTeachers(teachers_worksheet):
    teachers = {}
    teacher_row = 2
    teacher_col_map = map_headings(teachers_worksheet)
    while teachers_worksheet[teacher_col_map['Teacher ID'] + str(teacher_row)].value is not None:
        teacher_id = teachers_worksheet[teacher_col_map['Teacher ID'] + str(teacher_row)].value
        teachers[teacher_id] = {
            "name": teachers_worksheet[teacher_col_map['Teacher name'] + str(teacher_row)].value
        }
        teacher_row += 1
    return teachers

def readStudents(students_worksheet):
    students = {}
    student_row = 2
    student_col_map = map_headings(students_worksheet)
    while students_worksheet[student_col_map['Student ID'] + str(student_row)].value is not None:
        student_id = students_worksheet[student_col_map['Student ID'] + str(student_row)].value
        birthdate = students_worksheet[student_col_map['Date of Birth'] + str(student_row)].value
        students[student_id] = {
            "first_name": students_worksheet[student_col_map['First name'] + str(student_row)].value,
            "surname": students_worksheet[student_col_map['Surname'] + str(student_row)].value,
            "birth_date": {"dd": birthdate.day, "mm": birthdate.month, "yyyy": birthdate.year},
            "gender": students_worksheet[student_col_map['Gender'] + str(student_row)].value,
            "grade": str(students_worksheet[student_col_map['Grade'] + str(student_row)].value)
        }
        student_row += 1
    return students

def readClassroomAndAssets(excel_file):
    classroom_and_assets = {}
    w = load_workbook(excel_file)
    classroom_and_assets['classroom'], class_id = readClassroom(w['Classroom'])
    classroom_and_assets['asset'] = {
        class_id: {'teachers': readTeachers(w['Teachers']),
                   'students': readStudents(w['Students'])}
    }
    return classroom_and_assets

def writeJSON(outputfilename, classroom_and_assets):
    outputfile = open(outputfilename, 'w')
    outputfile.write(json.dumps(classroom_and_assets))
    outputfile.close()

if len(sys.argv) == 2:
    classroom_and_assets = readClassroomAndAssets(sys.argv[1])
    outputfilename = sys.argv[1][0:sys.argv[1].rfind('.')] + '.json'
    writeJSON(outputfilename, classroom_and_assets)
else:
    print('Make a classroom\n\nUsage: ' + sys.argv[0] + ' <excel filename>')
    print('\nOutput file will have same name as excel with .json extension\n')
