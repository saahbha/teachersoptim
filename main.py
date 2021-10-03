from flask import Flask, render_template, request, send_file
import xlsxwriter
from matcher import HardConstraintMatcher
from json import load
import pandas as pd

app = Flask(__name__)


@app.route('/')
def upload():
    return render_template("file_upload_form.html")


@app.route('/success', methods=['POST'])
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
        statTable = {}
        for line in output_stats:
            x = line.split(':')
            statTable[x[0]] = x[1]

        assignedStudents_values = {}
        assignedStudents_DF = pd.read_csv('Output_Assigned_Students.csv')
        assignedStudents_list = assignedStudents_DF.values.tolist()
        print(assignedStudents_list)
        for index, row in assignedStudents_DF.iterrows():
            assignedStudents_values[index] = assignedStudents_list[index]

        courses_values = {}
        courses_DF = pd.read_csv('Output_Courses.csv')
        courses_list = courses_DF.values.tolist()
        print(courses_list)
        for index, row in courses_DF.iterrows():
            courses_values[index] = courses_list[index]

        unassignedStudents_values = {}
        unassignedStudents_DF = pd.read_csv('Output_Unassigned_Students.csv')
        unassignedStudents_list = unassignedStudents_DF.values.tolist()
        print(unassignedStudents_list)
        for index, row in unassignedStudents_DF.iterrows():
            unassignedStudents_values[index] = unassignedStudents_list[index]

        nonuniqueStudents_values = {}
        nonuniqueStudents_DF = pd.read_csv('Output_BulletVoting_Students.csv')
        nonuniqueStudents_list = nonuniqueStudents_DF.values.tolist()
        print(nonuniqueStudents_list)
        for index, row in nonuniqueStudents_DF.iterrows():
            nonuniqueStudents_values[index] = nonuniqueStudents_list[index]

        return render_template("success.html",
                               statTable=statTable,
                               assignedStudents_columns=assignedStudents_DF.columns, assignedStudents_values=assignedStudents_values,
                               courses_columns=courses_DF.columns, courses_values=courses_values,
                               unassignedStudents_columns=unassignedStudents_DF.columns, unassignedStudents_values=unassignedStudents_values,
                               nonuniqueStudents_columns = nonuniqueStudents_DF.columns, nonuniqueStudents_values = nonuniqueStudents_values,
                               )


@app.route('/getCourseTemplate')  # this is a job for GET, not POST
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


@app.route('/getStudentTemplate')  # this is a job for GET, not POST
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

@app.route('/getAssignmentFile')  # this is a job for GET, not POST
def return_OutputAssignment():
    return send_file('Output_Assignment.xlsx', as_attachment=True)

@app.route('/getAssignedStudents')  # this is a job for GET, not POST
def return_AssignedStudents():
    return send_file('Output_Assigned_Students.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_Assigned_Students.csv',
                     as_attachment=True)


@app.route('/getCourses')  # this is a job for GET, not POST
def return_Courses():
    return send_file('Output_Courses.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_Courses.csv',
                     as_attachment=True)


@app.route('/getUnassignedStudents')  # this is a job for GET, not POST
def return_UnassignedStudents():
    return send_file('Output_Unassigned_Students.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_Unassigned_Students.csv',
                     as_attachment=True)

@app.route('/getBulletVotingStudents')  # this is a job for GET, not POST
def return_BulletVotingStudents():
    return send_file('Output_BulletVoting_Students.csv',
                     mimetype='text/csv',
                     attachment_filename='Output_BulletVoting_Students.csv',
                     as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
