import pyodbc, json

try:
    f = open('src/config.json', 'r')
    config = json.loads(f.read())
    database_path = config['database_path']
except IOError:
    print('Error: cant find config.json')
else:
    f.close()


def get_data(sql):
    try:
        con_string = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" \
                     f"DBQ={database_path}"
        conn = pyodbc.connect(con_string)
        print('Connected to database')

        cur = conn.cursor()
        cur.execute(sql)
        row = cur.fetchall()

    except pyodbc.Error as e:
        print("Error in connection", e)
        input('>> press enter key to exit...')
        exit()
    else:
        return row
