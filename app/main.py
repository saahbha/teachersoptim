from flask import Flask, render_template, request, send_file
import pandas as pd
from ortools.linear_solver import pywraplp
from pulp import *
import csv


app = Flask(__name__)  
 
@app.route('/')  
def upload():  
    return render_template("file_upload_form.html")  
 
@app.route('/success', methods = ['POST'])  
def success():  
    if request.method == 'POST':  
        f1 = request.files['file1']  
        f2 = request.files['file2']
        student_df = pd.read_excel(f1, sheet_name=0)
        courses_df = pd.read_excel(f2, sheet_name=1)

        def optim(studentsDataFile, coursesDataFile):
            # Read Column data
            courseNames = coursesDataFile['Course Name'].tolist()
            studentFirstName = studentsDataFile['first_name'].tolist()
            studentLastName = studentsDataFile['last_name'].tolist()

            studentFirstChoices = studentsDataFile['Preference 1'].tolist()
            studentSecondChoices = studentsDataFile['Preference 2'].tolist()
            studentThirdChoices = studentsDataFile['Preference 3'].tolist()
            courseMins = coursesDataFile['Test Min'].tolist()
            courseMaxs = coursesDataFile['Test Max'].tolist()
            S = len(studentFirstChoices)
            C = len(courseMins)

            model = LpProblem("TAS Matching", LpMaximize)
            solver = getSolver('PULP_CBC_CMD')

            #variables
            studentAssignments = [[LpVariable("S%dC%d"%(s,c), 0, 1, LpInteger) 
                                for c in range(C)] for s in range(S)]
            classWillRun = [LpVariable("C%dr"%c, 0, 1, LpInteger) 
                                for c in range(C)]

            #Sum of students in each class
            sumStudentsInClass = [LpAffineExpression() for c in range(C)]
            for c in range(C):
                for s in range(S):
                    sumStudentsInClass[c] += studentAssignments[s][c]

            #Constraints described by Dr. Miller in email
            runConstraints = [[LpConstraint() for s in range(S)] for c in range(C)]
            for c in range(C):
                for s in range(S):  
                    constraint = studentAssignments[s][c] - classWillRun[c]
                    runConstraints[c][s] = LpConstraint(e=constraint, sense=-1, name="C%dS%dr"%(c, s), rhs=0)
            sizeConstraints = [LpConstraint() for c in range(C)]    
            for c in range(C):
                constraint = classWillRun[c] - sumStudentsInClass[c]
                sizeConstraints[c] = LpConstraint(e=constraint, sense=-1, name="C%ds"%c, rhs=0)

            #Constraints for each class
            classConstraints = [[LpConstraint() for c in range(C)] for i in range(2)]
            for c in range(C):    
                hardMin = sumStudentsInClass[c] - courseMins[c]*classWillRun[c]
                hardMax = sumStudentsInClass[c] - courseMaxs[c]*classWillRun[c]
                cMin = LpConstraint(e=hardMin, sense=1, name="C%dm"%c, rhs=0)
                cMax = LpConstraint(e=hardMax, sense=-1, name="C%dM"%c, rhs=0)
                classConstraints[0][c] = cMin
                classConstraints[1][c] = cMax

            #Constraint limiting the number of classes a student should be assigned to
            maxAssignmentConstraint = [LpConstraint() for s in range(S)]
            for s in range(S):
                sumOfAssignments = LpAffineExpression()
                for c in range(C):
                    sumOfAssignments += studentAssignments[s][c]
                maxAssignmentConstraint[s] = sumOfAssignments <= 1
                
            #Student preference vector (for the objective function)
            # studentPreferences = [[0 for c in range(C)] for s in range(S)]
            # for s in range(S):
            #     c1 = (studentFirstChoices[s])-1
            #     c2 = (studentSecondChoices[s])-1
            #     c3 = (studentThirdChoices[s])-1
            #     #students who bullet vote will have their single preference value be 1
            #     studentPreferences[s][c1] = 5
            #     studentPreferences[s][c2] = 3   
            #     studentPreferences[s][c3] = 1
            
            studentPreferences = [[0 for c in range(C)] for s in range(S)]
            for s in range(S):
                coursePreferences = []
                for c in range(C):
                    if (studentFirstChoices[s] == (c + 1)):
                        coursePreferences.append(5)
                    elif (studentSecondChoices[s] == (c + 1)):
                        coursePreferences.append(3)
                    elif (studentThirdChoices[s] == (c + 1)):
                        coursePreferences.append(1)
                    else:
                        coursePreferences.append(0)
                studentPreferences[s] = coursePreferences
            
            #Objective function
            objective = LpAffineExpression()
            for c in range(C):
                for s in range(S):
                    objective += studentPreferences[s][c]*studentAssignments[s][c]
                #objective += (5*S/C)*classWillRun[c]
                
            #send data to model
            model += objective
            for s in range(S):
                model += maxAssignmentConstraint[s]
            for c in range(C):
                for s in range(S):
                    model += runConstraints[c][s]
                model += sizeConstraints[c]
                for i in range(len(classConstraints)):
                    model += classConstraints[i][c]
            model.writeMPS("TAS.mps")

            #Solve
            model.writeLP("TAS.lp")
            model.solve(solver)
            print("Status:", LpStatus[model.status])
            print("Objective value: ", value(model.objective))

            #Convert the variables to their values
            for c in range(C):
                for s in range(S):    
                    studentAssignments[s][c] = int(studentAssignments[s][c].varValue)

            #Output Results
            numFirstChoiceAssignment = 0
            numSecondChoiceAssignment = 0
            numThirdChoiceAssignment = 0
            numNoChoiceAssignment = 0
            numNoAssignment = 0
            numMultiAssignment = 0
            studentIdAssignments = [[-1] for s in range(S)]
            courseSizes = [0 for c in range(C)]
            courseFirstChoices = [0 for c in range(C)]
            courseSecondChoices = [0 for c in range(C)]
            courseThirdChoices = [0 for c in range(C)]
            for s in range(S):
                firstChoice = studentFirstChoices[s]
                secondChoice = studentSecondChoices[s]
                thirdChoice = studentThirdChoices[s]
                sumAssignments = 0
                for c in range(C):        
                    #Courses Stats
                    courseId = c+1
                    if (courseId == firstChoice): 
                        courseFirstChoices[c] += 1
                    elif (courseId == secondChoice):
                        courseSecondChoices[c] += 1
                    elif (courseId == thirdChoice):
                        courseThirdChoices[c] += 1  
                    if (studentAssignments[s][c] == 1): 
                        courseSizes[c] += 1
                        
                    #Students Stats
                    if (studentAssignments[s][c] == 1): 
                        if (courseId == firstChoice): 
                            numFirstChoiceAssignment+=1
                        elif (courseId == secondChoice):
                            numSecondChoiceAssignment+=1
                        elif (courseId == thirdChoice):
                            numThirdChoiceAssignment+=1
                        else:
                            numNoChoiceAssignment+=1
                        sumAssignments+=1
                        if studentIdAssignments[s][0] == -1:
                            studentIdAssignments[s][0] = courseId
                        else:
                            studentIdAssignments[s].append(courseId)
                if sumAssignments > 1:
                    numMultiAssignment+=1
                elif sumAssignments == 0:
                    numNoAssignment+=1
                    
            #Print some stats            
            prefobj = 5*numFirstChoiceAssignment+3*numSecondChoiceAssignment+numThirdChoiceAssignment
            print('Placement Objective Value: %d' % prefobj)
            print('First Choice Assignments: %.5f (%d/%d)' 
                % (float(numFirstChoiceAssignment)/S, numFirstChoiceAssignment, S))
            print('Second Choice Assignments: %.5f (%d/%d)' 
                % (float(numSecondChoiceAssignment)/S, numSecondChoiceAssignment, S))
            print('Third Choice Assignments: %.5f (%d/%d)'
                % (float(numThirdChoiceAssignment)/S, numThirdChoiceAssignment, S))
            print('No Choice Assignments: %.5f (%d/%d)'
                % (float(numNoChoiceAssignment)/S, numNoChoiceAssignment, S))
            print('Multi Assignments: %.5f (%d/%d)'
                % (float(numMultiAssignment)/S, numMultiAssignment, S))
            print('No Assignments: %.5f (%d/%d)'
                % (float(numNoAssignment)/S, numNoAssignment, S))


            #Output to CSV
            with open('Output_Courses.csv', mode='w') as course_file:
                course_writter = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
                #write column names
                columnNames = []
                columnNames.append('Course number')
                columnNames.append('Course name')
                columnNames.append('First choice')
                columnNames.append('Second choice')
                columnNames.append('Third choice')
                columnNames.append('Weight')
                columnNames.append('Minimum class size')
                columnNames.append('Maximum class size')
                columnNames.append('Students assigned')
                course_writter.writerow(columnNames)
                
                #write row data
                for c in range(C):
                    rowData = []
                    c1 = courseFirstChoices[c]
                    c2 = courseSecondChoices[c]
                    c3 = courseThirdChoices[c]
                    
                    rowData.append(c+1) #'Course number'
                    rowData.append(courseNames[c]) #'Course name'
                    rowData.append(c1) #'First choice'
                    rowData.append(c2) #'Second choice'
                    rowData.append(c3) #'Third choice'
                    rowData.append(5*c1+3*c2+c3) #'Weight'
                    rowData.append(courseMins[c]) #'Minimum class size'
                    rowData.append(courseMaxs[c]) #'Maximum class size'
                    rowData.append(courseSizes[c]) #'Students assigned'
                    course_writter.writerow(rowData)

            with open('Output_Assigned_Students.csv', mode='w') as course_file:
                student_writer = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
                #write column names
                columnNames = []
                columnNames.append('Student ID')
                columnNames.append('First Name')
                columnNames.append('Last Name')
                columnNames.append('Course Assignment')
                student_writer.writerow(columnNames)
                
                #write row data
                for s in range(S):
                    if (studentIdAssignments[s][0] != -1):
                        ca = "%d"%studentIdAssignments[s][0]
                        for i in range(1, len(studentIdAssignments[s])):
                            ca += ", %d"%studentIdAssignments[s][i]
                        
                        rowData = []
                        rowData.append(s+1)#'Student ID'
                        rowData.append(studentFirstName[s])#'First Name'
                        rowData.append(studentLastName[s])#'Last Name'
                        rowData.append(ca)#'Course Assignment'
                        student_writer.writerow(rowData)
                    
            with open('Output_Unassigned_Students.csv', mode='w') as course_file:
                student_writer = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
                #write column names
                columnNames = []
                columnNames.append('Student ID')
                columnNames.append('First Name')
                columnNames.append('Last Name')
                student_writer.writerow(columnNames)
                
                #write row data
                for s in range(S):
                    if (studentIdAssignments[s][0] == -1):            
                        rowData = []
                        rowData.append(s+1)#'Student ID'
                        rowData.append(studentFirstName[s])#'First Name'
                        rowData.append(studentLastName[s])#'Last Name'
                        student_writer.writerow(rowData)
                        
            with open('Output_Stats.csv', mode='w') as course_file:
                stats_writer = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
                #write column names
                stats_writer.writerow(['First Choice Assignments: %.5f (%d/%d)' 
                    % (float(numFirstChoiceAssignment)/S, numFirstChoiceAssignment, S)])
                stats_writer.writerow(['Second Choice Assignments: %.5f (%d/%d)' 
                    % (float(numSecondChoiceAssignment)/S, numSecondChoiceAssignment, S)])
                stats_writer.writerow(['Third Choice Assignments: %.5f (%d/%d)'
                    % (float(numThirdChoiceAssignment)/S, numThirdChoiceAssignment, S)])
                stats_writer.writerow(['No Choice Assignments: %.5f (%d/%d)'
                    % (float(numNoChoiceAssignment)/S, numNoChoiceAssignment, S)])
                stats_writer.writerow(['Multi Assignments: %.5f (%d/%d)'
                    % (float(numMultiAssignment)/S, numMultiAssignment, S)])
                stats_writer.writerow(['No Assignments: %.5f (%d/%d)'
                    % (float(numNoAssignment)/S, numNoAssignment, S)])
        

        optim(student_df, courses_df)


        # open csv file
        a = open("Output_Stats.csv", 'r')
        a = a.readlines()
        return render_template("sucess.html", name = a)  


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

if __name__ == '__main__':  
    app.run(debug = True)  