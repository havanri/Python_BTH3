import mysql.connector
import pandas

excel_data_df = pandas.read_excel('D:\\input2.xlsx', sheet_name = 'MAU', header = 10, usecols = range(1, 8), nrows=52)
print(excel_data_df)

mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="student_python"
)
mycursor = mydb.cursor()
for row in excel_data_df.iterrows():
    student = row[1].to_dict()

    ini_Key = ['0', '1', '2', '3', '4', '5']

    student_Item = dict(zip(ini_Key, list(student.values())))
    sql = "INSERT INTO students(firstname, lastname, birth, dToan, dLy, dHoa) VALUES (%s, %s, %s, %s, %s, %s)"
    val = (student_Item['0'], student_Item['1'], student_Item['2'], student_Item['3'], student_Item['4'], student_Item['5'])
    mycursor.execute(sql, val)
    print(mycursor.rowcount, "record inserted.")
    mydb.commit()
mydb.close()
