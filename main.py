from flask import Flask, render_template, request, send_file, send_from_directory
import xlsxwriter
from matcher import HardConstraintMatcher
from json import load
import pandas as pd
from csv import writer, QUOTE_MINIMAL

app = Flask(__name__)  
 
@app.route('/')  
def upload():  
    return render_template("file_upload_form.html")  
 
@app.route('/success', methods = ['POST'])  
def success():  
    if request.method == 'POST':
        # Solve from input
        students_FileLocation = request.files['file1']
        courses_FileLocation = request.files['file2']
        with open('config.json') as config_file:
            config = load(config_file)
            matcher = HardConstraintMatcher(
                students_FileLocation=students_FileLocation,
                students_SheetName=config["students_SheetName"],
                students_Columns=config["students_Columns"],
                courses_FileLocation=courses_FileLocation,
                courses_SheetName=config["courses_SheetName"],
                courses_Columns=config["courses_Columns"]
            )
        matcher.solve()  # this is the step that takes a long time
        matcher.outputResults()

        # open csv file
        a = open("Output_Stats.csv", 'r')
        a = a.readlines()
        output_stats = [val.replace('\n', '') for val in a]
        result = {}
        for line in output_stats:
            x = line.split(':')
            result[x[0]] = x[1]

        preview = {}
        df = pd.read_csv('Output_Assigned_Students.csv')
        df_list = df.values.tolist()
        print(df_list)
        for index, row in df.iterrows():
            preview[index] = df_list[index]

        preview2 = {}
        df2 = pd.read_csv('Output_Courses.csv')
        df_list2 = df2.values.tolist()
        for i, row in df2.iterrows():
            preview2[i] = df_list2[i]

        preview3 = {}
        df3 = pd.read_csv('Output_Unassigned_Students.csv')
        df_list3 = df3.values.tolist()
        print(df_list3)
        for index, row in df3.iterrows():
            preview3[index] = df_list3[index]

        return render_template("sucess.html", name=a, val=preview, jv=result, names=df.columns, names2=df2.columns, val2=preview2, names3=df3.columns, val3=preview3)


@app.route('/getCourseTemplate') # this is a job for GET, not POST
def plot_CourseTemplate():
    columns = {}
    with open('config.json') as config_file:
        config = load(config_file)
        columns = config["courses_Columns"]

    # write column names
    columnNames = []
    columnNames.append(columns["Name"])
    columnNames.append(columns["Min"])
    columnNames.append(columns["Max"])

    workbook = xlsxwriter.Workbook('courses_template.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write_row('A1', columnNames)
    workbook.close()

    return send_file('courses_template.xlsx', as_attachment=True)

@app.route('/getStudentTemplate') # this is a job for GET, not POST
def plot_StudentTemplate():
    columns = {}
    with open('config.json') as config_file:
        config = load(config_file)
        columns = config["students_Columns"]

    # write column names
    columnNames = []
    columnNames.append(columns["First_Name"])
    columnNames.append(columns["Last_Name"])
    columnNames.append(columns["P1"])
    columnNames.append(columns["P2"])
    columnNames.append(columns["P3"])

    workbook = xlsxwriter.Workbook('students_template.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write_row('A1', columnNames)
    workbook.close()

    return send_file('students_template.xlsx', as_attachment=True)

@app.route('/getPlotCSV') # this is a job for GET, not POST
def plot_csv():

    return send_file('Output_Assigned_Students.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_Assigned_Students.csv',
                     as_attachment=True)

@app.route('/getPlotCSV1') # this is a job for GET, not POST
def plot_csv1():
    return send_file('Output_Courses.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_Courses.csv',
                     as_attachment=True)

@app.route('/getPlotCSV2') # this is a job for GET, not POST
def plot_csv2():
    return send_file('Output_Unassigned_Students.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_Unassigned_Students.csv',
                     as_attachment=True)

if __name__ == '__main__':  
    app.run(debug = True)  