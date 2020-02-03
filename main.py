from flask import Flask, render_template, request, redirect
import sys
import os
import xlrd
from flaskext.mysql import MySQL
from werkzeug.utils import secure_filename
import mysql.connector
from mysql.connector import Error
from flask_cors import CORS

connection = mysql.connector.connect(host='localhost', database='myexceldata', user='root', password='')
cursor = connection.cursor(buffered=True)


mysql = MySQL()
app = Flask(__name__)
message = ''

mysql.init_app(app)

UPLOAD_FOLDER = './'
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(os.path.join(app.instance_path, 'excel_data'), exist_ok=True)


CORS(app)

# upload the file
@app.route('/uploads')
def uploads():
    return render_template('upload.html')


# put data from uploaded excel into db
@app.route('/uploadsuccess', methods=['POST'])
def success():
    try:
        if request.method == 'POST':
            f = request.files['upload']
            # get the file name and save it to a directory in folder "excel_data"
            filename = secure_filename(f.filename)
            f.save(os.path.join(app.config['UPLOAD_FOLDER'], 'excel_data' , secure_filename(f.filename)))
            # Reading an excel file using Python 
            # Give the location of the file 
            loc= os.path.join(app.config['UPLOAD_FOLDER'], 'excel_data' , secure_filename(f.filename))
            # To open Workbook
            wb = xlrd.open_workbook(loc) 
            sheet = wb.sheet_by_index(0) 
            val = []
            tempArr = []
            # get number of rows and columns
            numrows = sheet.nrows
            numcolumns = sheet.ncols
            dbdata= sheet.row_values(1) 
            for x in range(1,numrows):
              
                  column2=str(int(sheet.cell_value(x , 1))) 
                  tempArr = [sheet.cell_value(x , 0),column2]
                  val.append(tempArr)
            sql = """INSERT INTO map 
                   (column1, column2) 
                   VALUES ( %s, %s )"""      
            cursor.executemany(sql, val)
            connection.commit() 
    except Exception as e:
        print(e)    
    return render_template('success.html', val= val)

if __name__ == '__main__':
    app.secret_key = os.urandom(12)
    app.run(debug=True)
    