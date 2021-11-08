import json
import mymodule, database


def main_menu():
    print('**********************ENROLLMENT REPORTS**********************')
    print('[0] GENERATE REPORT BY SECTION\n[1] GENERATE REPORT BY SEX\n[2] GENERATE STUDENT LISTS BY YEAR')
    ans = input('>> Please select from menu:\n>> ')

    if ans == '0':
        data = database.get_data(
            'SELECT COURSE, YEAR1, SECTION1, COUNT(STUDE_NO) FROM student GROUP BY COURSE, YEAR1, SECTION1')
        res = enrollment.create_report_by_section(data)
    elif ans == '1':
        data = database.get_data(
            'SELECT COURSE, YEAR1, SEX, COUNT(SEX) FROM student WHERE SEX <> null GROUP BY COURSE, YEAR1, SEX')
        res = enrollment.create_report_by_sex(data)
    elif ans == '2':

        data = database.get_data(
            "SELECT COURSE, YEAR1, STUDNOLOC, FNAME, MNAME, SNAME, SEX, STATUS, "
            "TIME1, TIME2, TIME3, TIME4, TIME5, TIME6, TIME7, TIME8, TIME9, TIME10, "
            "UNIT1, UNIT2, UNIT3, UNIT4, UNIT5, UNIT6, UNIT7, UNIT8, UNIT9, UNIT10, "
            "SUBJECT1, SUBJECT2, SUBJECT3, SUBJECT4, SUBJECT5, SUBJECT6, SUBJECT7, SUBJECT8, SUBJECT9, SUBJECT10 "
            "FROM student ORDER BY COURSE, YEAR1, FNAME")
        enrollment.create_student_list(data)
    else:
        print('Invalid input, please try again.')
        main_menu()

    if res:
        print('>> Report successfully created!')
        input('>> press enter key to exit...')
        exit()


try:
    f = open('src/config.json', 'r')
    config = json.loads(f.read())
    semester = config['semester']
    database_path = config['database_path']
    enrollment = mymodule.EnrollmentReport(config)
    main_menu()
except IOError:
    print('Error: cant find config.json')
finally:
    f.close()
