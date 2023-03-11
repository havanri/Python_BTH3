import mysql.connector
import pandas
from datetime import datetime

excel_data_df = pandas.read_excel('D:\\input2.xlsx', sheet_name = 'MAU', header = 10, usecols = range(1, 8), nrows=52)

mydb = mysql.connector.connect(
    host = "localhost",
    user = "root",
    password = "",
    database = "student_python"
)
mycursor = mydb.cursor()
for row in excel_data_df.iterrows():
    student = row[1].to_dict()

    ini_Key = ['0', '1', '2', '3', '4', '5', '6']
    student_Item = dict(zip(ini_Key, list(student.values())))

    sql = "INSERT INTO students(firstname, lastname, birth, dToan, dLy, dHoa) VALUES(%s, %s, %s, %s, %s, %s)"
    val = (student_Item['1'], student_Item['2'], datetime.strptime(student_Item['3'], '%d/%m/%Y'), student_Item['4'], student_Item['5'], student_Item['6'])
    mycursor.execute(sql, val)

    mydb.commit()
mydb.close()
