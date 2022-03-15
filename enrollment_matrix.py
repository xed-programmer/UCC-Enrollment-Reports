import json
import mymodule, database


def main_menu():
    print('**********************ENROLLMENT REPORTS**********************')
    print('[0] GENERATE REPORT BY SECTION\n[1] GENERATE REPORT BY SEX\n[2] GENERATE STUDENT LISTS BY YEAR')
    ans = input('>> Please select from menu:\n>> ')

    if ans == '0':
        table_name = input('Enter table name: ')
        data = database.get_data(
            'SELECT COURSE, YEAR1, SECTION1, COUNT(STUDE_NO) FROM ' + table_name + ' GROUP BY COURSE, YEAR1, SECTION1')
        print(len(data))
        res = enrollment.create_report_by_section(data)

    elif ans == '1':
        table_name = input('Enter table name: ')
        data = database.get_data(
            'SELECT COURSE, YEAR1, SEX, COUNT(SEX) FROM ' + table_name + ' WHERE SEX <> null GROUP BY COURSE, YEAR1, SEX')
        print(len(data))
        res = enrollment.create_report_by_sex(data)

    elif ans == '2':
        table_name = input('Enter table name: ')
        data = database.get_data(
            "SELECT COURSE, YEAR1, STUDNOLOC, FNAME, MNAME, SNAME, SEX, STATUS, "
            "TIME1, TIME2, TIME3, TIME4, TIME5, TIME6, TIME7, TIME8, TIME9, TIME10, "
            "UNIT1, UNIT2, UNIT3, UNIT4, UNIT5, UNIT6, UNIT7, UNIT8, UNIT9, UNIT10, "
            "SUBJECT1, SUBJECT2, SUBJECT3, SUBJECT4, SUBJECT5, SUBJECT6, SUBJECT7, SUBJECT8, SUBJECT9, SUBJECT10 "
            "FROM " + table_name + " ORDER BY COURSE, YEAR1, FNAME")
        print(len(data))
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
    input('>> press enter key to exit...')
    exit()
finally:
    f.close()
