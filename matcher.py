''' 
What does a matcher do?
    Takes 2 files for input and outputs 4 files
        -takes input describing students and courses
            -Whats the input data we need to solve the problem?
                -For making a solution, we need:
                    -student First Choices
                    -student Second Choices
                    -student Third Choices
                    -course Mins
                    -course Maxs
                -For outputting results formatted, we need the above plus:
                    -student first name
                    -student last name
                    -student id
                    -course name
                    -course id
        -use the input data to come up with a solution
        -outputs stats, course data w/ sizes, student data w/ assignments, student unassignments

    What are the steps for a matcher in general???
        -take in input
        -solve
        -output results
        
    What is common for all matchers???
        -inputs are all the same format
        -each matcher has to match/solve the problem
        -outputs are all the same format

    What are the steps for a matcher based on PuLP linear programming???
        -read from excel files for students and courses and put columns into arrays/lists
        
        -solve the model
            -initialize a model and a solver
            -initialize variables
            -create constraints from the variables and input excel data
            -create an objective equation from the variables and input excel data
            -add this the constraints and objective to the model
            -use the solver to solve the model
        
        -output results
    
    Whats the difference between the hard constraints matcher and a soft constraints matcher?
        -initialize two sets of penalty variables corrisponding to each course
        -add a penalty variable to each course/class constraints
        -add a penalty term to the objective for each penalty variable
'''

from pandas import read_excel
from csv import writer, QUOTE_MINIMAL
import xlsxwriter
from pulp import LpProblem, LpMaximize, getSolver, LpVariable, LpInteger, LpAffineExpression, LpConstraint, LpStatus, \
    value


class Matcher:
    """
    Example Usage:
        from matcher import HardConstraintMatcher
        
        matcher = HardConstraintMatcher(
            students_FileLocation = 'Students.xlsx',
            students_SheetName = 0,
            students_Columns = {
                "First_Name": "First Name",
                "Last_Name": "Last Name",
                "P1": "Preference 1",
                "P2": "Preference 2",
                "P3": "Preference 3",
            },
            courses_FileLocation = 'Courses.xlsx',
            courses_SheetName = 0,
            courses_Columns = {
                "Name": "Course Name",
                "Min": "Minimum Size",
                "Max": "Maximum Size"
            }
        )
        matcher.solve() #this is the step that takes a long time
        matcher.outputResults()
    """

    def __init__(self,
                 # Default Inputs
                 students_FileLocation='data\MOCK_Students.xlsx',
                 students_SheetName=0,
                 students_Columns={
                     "First_Name": "first_name",
                     "Last_Name": "last_name",
                     "P1": "Preference 1",
                     "P2": "Preference 2",
                     "P3": "Preference 3",
                 },
                 courses_FileLocation='data\MOCK_Courses.xlsx',
                 courses_SheetName=1,
                 courses_Columns={
                     "Name": "Course Name",
                     "Min": "Test Min",
                     "Max": "Test Max"
                 },
                 p1_ObjectiveValue=5,
                 p2_ObjectiveValue=3,
                 p3_ObjectiveValue=1
                 ):
        self.c1Value = p1_ObjectiveValue
        self.c2Value = p2_ObjectiveValue
        self.c3Value = p3_ObjectiveValue

        # Open Excel File
        studentsData = read_excel(students_FileLocation,
                                  sheet_name=students_SheetName)
        coursesData = read_excel(courses_FileLocation,
                                 sheet_name=courses_SheetName)

        # Read Column data
        self.studentFirstName = studentsData[students_Columns["First_Name"]].tolist()
        self.studentLastName = studentsData[students_Columns["Last_Name"]].tolist()
        self.courseNames = coursesData[courses_Columns["Name"]].tolist()
        self.S = len(self.studentFirstName)
        self.C = len(self.courseNames)

        self.studentFirstChoices = studentsData[students_Columns["P1"]].tolist()
        self.studentSecondChoices = studentsData[students_Columns["P2"]].tolist()
        self.studentThirdChoices = studentsData[students_Columns["P3"]].tolist()
        self.courseMins = coursesData[courses_Columns["Min"]].tolist()
        self.courseMaxs = coursesData[courses_Columns["Max"]].tolist()

    @staticmethod
    def notUniqueChoices(c1, c2, c3):
        return c1 == c2 or c1 == c3 or c2 == c3 or (not isinstance(c1, int)) or (not isinstance(c2, int)) or (
            not isinstance(c3, int))

    ##### The folowing 3 functions should be expanded upon for each type of Matcher #####
    def initVariables(self):
        self.studentAssignments = [[LpVariable("S%dC%d" % (s, c), 0, 1, LpInteger)
                                    for c in range(self.C)] for s in range(self.S)]

    def makeObjective(self):
        # Student preference vector (for the objective function)
        self.studentPreferences = [[0 for c in range(self.C)] for s in range(self.S)]

        # for s in range(self.S):
        #     coursePreferences = []
        #     for c in range(self.C):
        #         if (self.studentFirstChoices[s] == (c + 1)):
        #             coursePreferences.append(5)
        #         elif (self.studentSecondChoices[s] == (c + 1)):
        #             coursePreferences.append(3)
        #         elif (self.studentThirdChoices[s] == (c + 1)):
        #             coursePreferences.append(1)
        #         else:
        #             coursePreferences.append(0)
        #     self.studentPreferences[s] = coursePreferences
        for s in range(self.S):
            c1 = self.studentFirstChoices[s]
            c2 = self.studentSecondChoices[s]
            c3 = self.studentThirdChoices[s]

            uniqueChoices = 0 if self.notUniqueChoices(c1, c2, c3) else 1

            if isinstance(c1, int):
                self.studentPreferences[s][c1 - 1] = (self.c1Value - 1) * uniqueChoices + 1
            if isinstance(c2, int):
                self.studentPreferences[s][c2 - 1] = (self.c2Value - 1) * uniqueChoices + 1
            if isinstance(c3, int):
                self.studentPreferences[s][c3 - 1] = (self.c3Value - 1) * uniqueChoices + 1

    def makeConstraints(self):
        self.sumStudentsInClass = [LpAffineExpression() for c in range(self.C)]
        for c in range(self.C):
            for s in range(self.S):
                self.sumStudentsInClass[c] += self.studentAssignments[s][c]

    def initProblem(self):
        self.model = LpProblem("TAS Matching", LpMaximize)
        self.initVariables()
        self.makeObjective()
        self.makeConstraints()

    def solve(self, solver="PULP_CBC_CMD"):
        self.initProblem()
        # self.model.writeLP("TAS.lp")
        self.model.solve(getSolver(solver))
        print("Status:", LpStatus[self.model.status])
        print("Objective value: ", value(self.model.objective))

    def outputResults(self):
        # Convert the variables to their values
        for c in range(self.C):
            for s in range(self.S):
                self.studentAssignments[s][c] = int(self.studentAssignments[s][c].varValue)

        # Output Results
        numFirstChoiceAssignment = 0
        numSecondChoiceAssignment = 0
        numThirdChoiceAssignment = 0
        numNoChoiceAssignment = 0
        numNoAssignment = 0
        numMultiAssignment = 0
        self.studentIdAssignments = [[-1] for s in range(self.S)]
        courseSizes = [0 for c in range(self.C)]
        courseFirstChoices = [0 for c in range(self.C)]
        courseSecondChoices = [0 for c in range(self.C)]
        courseThirdChoices = [0 for c in range(self.C)]
        courseAssignedFirstChoices = [0 for c in range(self.C)]
        courseAssignedSecondChoices = [0 for c in range(self.C)]
        courseAssignedThirdChoices = [0 for c in range(self.C)]
        for s in range(self.S):
            firstChoice = self.studentFirstChoices[s]
            secondChoice = self.studentSecondChoices[s]
            thirdChoice = self.studentThirdChoices[s]
            sumAssignments = 0
            for c in range(self.C):
                # Courses Stats
                courseId = c + 1
                # total choice and assigned choice
                if (courseId == firstChoice):
                    courseFirstChoices[c] += 1
                    if (self.studentAssignments[s][c] == 1):
                        courseAssignedFirstChoices[c] += 1
                elif (courseId == secondChoice):
                    courseSecondChoices[c] += 1
                    if (self.studentAssignments[s][c] == 1):
                        courseAssignedSecondChoices[c] += 1
                elif (courseId == thirdChoice):
                    courseThirdChoices[c] += 1
                    if (self.studentAssignments[s][c] == 1):
                        courseAssignedThirdChoices[c] += 1
                if (self.studentAssignments[s][c] == 1):
                    courseSizes[c] += 1

                # Students Stats
                if (self.studentAssignments[s][c] == 1):
                    if (courseId == firstChoice):
                        numFirstChoiceAssignment += 1
                    elif (courseId == secondChoice):
                        numSecondChoiceAssignment += 1
                    elif (courseId == thirdChoice):
                        numThirdChoiceAssignment += 1
                    else:
                        numNoChoiceAssignment += 1
                    sumAssignments += 1
                    if self.studentIdAssignments[s][0] == -1:
                        self.studentIdAssignments[s][0] = courseId
                    else:
                        self.studentIdAssignments[s].append(courseId)
            if sumAssignments > 1:
                numMultiAssignment += 1
            elif sumAssignments == 0:
                numNoAssignment += 1

        # Print some stats
        prefobj = self.c1Value * numFirstChoiceAssignment + self.c2Value * numSecondChoiceAssignment + self.c3Value * numThirdChoiceAssignment
        totalWeightStat = ('Total Course Weight Achieved: %.5f (%d/%d)'
              % (float(prefobj) / (self.S*self.c1Value), prefobj, self.S*self.c1Value))
        firstChoiceStat = ('First Choice Assignments: %.5f (%d/%d)'
              % (float(numFirstChoiceAssignment) / self.S, numFirstChoiceAssignment, self.S))
        secondChoiceStat = ('Second Choice Assignments: %.5f (%d/%d)'
              % (float(numSecondChoiceAssignment) / self.S, numSecondChoiceAssignment, self.S))
        thirdChoiceStat = ('Third Choice Assignments: %.5f (%d/%d)'
              % (float(numThirdChoiceAssignment) / self.S, numThirdChoiceAssignment, self.S))
        noChoiceStat = ('No Choice Assignments: %.5f (%d/%d)'
              % (float(numNoChoiceAssignment) / self.S, numNoChoiceAssignment, self.S))
        multiChoiceStat = ('Multi Assignments: %.5f (%d/%d)'
              % (float(numMultiAssignment) / self.S, numMultiAssignment, self.S))
        noAssignmentStat = ('No Assignments: %.5f (%d/%d)'
              % (float(numNoAssignment) / self.S, numNoAssignment, self.S))
        print(totalWeightStat)
        print(firstChoiceStat)
        print(secondChoiceStat)
        print(thirdChoiceStat)
        print(noChoiceStat)
        print(multiChoiceStat)
        print(noAssignmentStat)

        # Output to xlsx
        workbook = xlsxwriter.Workbook('Output_Assignment.xlsx', {'strings_to_numbers':  True})
        workbook.add_worksheet("Courses")
        workbook.add_worksheet("Student Assignments")
        workbook.add_worksheet("Student Unassignments")
        workbook.add_worksheet("Stats")

        # Add a format. Light red fill with dark red text.
        highlight_red = workbook.add_format({'bg_color': '#FFC7CE',
                                       'font_color': '#9C0006'})
        # Add a format. Yellow fill with dark yellow text.
        highlight_yellow = workbook.add_format({'bg_color': '#FFEB9C',
                                               'font_color': '#9C5700'})
        # Add a format. Green fill with dark green text.
        highlight_green = workbook.add_format({'bg_color': '#C6EFCE',
                                       'font_color': '#006100'})
        text_red = workbook.add_format({'font_color': '#FF0000'})

        # Output to CSV
        with open('Output_Courses.csv', mode='w', encoding='utf-8') as course_file:
            course_writter = writer(course_file, delimiter=',', quotechar='"', quoting=QUOTE_MINIMAL,
                                    lineterminator='\n')
            worksheet = workbook.get_worksheet_by_name("Courses")

            # write column names
            columnNames = []
            columnNames.append('Course number')
            columnNames.append('Course name')
            columnNames.append('Total Choices')  # 'Total Choices'
            columnNames.append('Total First choices')
            columnNames.append('Total Second choices')
            columnNames.append('Total Third choices')
            columnNames.append('Weight')
            columnNames.append('Minimum class size')
            columnNames.append('Maximum class size')
            columnNames.append('Students assigned')
            columnNames.append('Assigned First choices')
            columnNames.append('Assigned Second choices')
            columnNames.append('Assigned Third choices')
            columnNames.append('Assigned Weight')
            course_writter.writerow(columnNames)
            worksheet.write_row('A1', columnNames)

            # write row data
            for c in range(self.C):
                rowData = []
                c1 = courseFirstChoices[c]
                c2 = courseSecondChoices[c]
                c3 = courseThirdChoices[c]
                ca1 = courseAssignedFirstChoices[c]
                ca2 = courseAssignedSecondChoices[c]
                ca3 = courseAssignedThirdChoices[c]

                rowData.append(c + 1)  # 'Course number'
                rowData.append(self.courseNames[c])  # 'Course name'
                rowData.append(c1 + c2 + c3)  # 'Total Choices'
                rowData.append(c1)  # 'Total First choices'
                rowData.append(c2)  # 'Total Second choices'
                rowData.append(c3)  # 'Total Third choices'
                rowData.append(self.c1Value * c1 + self.c2Value * c2 + self.c3Value * c3)  # 'Weight'
                rowData.append(self.courseMins[c])  # 'Minimum class size'
                rowData.append(self.courseMaxs[c])  # 'Maximum class size'
                rowData.append(courseSizes[c])  # 'Students assigned'
                rowData.append(ca1)  # 'Assigned First choices'
                rowData.append(ca2)  # 'Assigned Second choices'
                rowData.append(ca3)  # 'Assigned Third choices'
                rowData.append(self.c1Value * ca1 + self.c2Value * ca2 + self.c3Value * ca3)  # 'Assigned Weight'
                course_writter.writerow(rowData)
                worksheet.write_row(c+1, 0, rowData)

        with open('Output_Assigned_Students.csv', mode='w', encoding='utf-8') as course_file:
            student_writer = writer(course_file, delimiter=',', quotechar='"', quoting=QUOTE_MINIMAL,
                                    lineterminator='\n')
            worksheet = workbook.get_worksheet_by_name("Student Assignments")

            # write column names
            columnNames = []
            columnNames.append('Student ID')
            columnNames.append('First Name')
            columnNames.append('Last Name')
            columnNames.append('First choice')
            columnNames.append('Second choice')
            columnNames.append('Third choice')
            columnNames.append('Course Assignment')
            student_writer.writerow(columnNames)
            worksheet.write_row('A1', columnNames)

            # write row data
            row = 1
            for s in range(self.S):
                if (self.studentIdAssignments[s][0] != -1):
                    ca = "%d" % self.studentIdAssignments[s][0]
                    for i in range(1, len(self.studentIdAssignments[s])):
                        ca += ", %d" % self.studentIdAssignments[s][i]

                    rowData = []
                    rowData.append(s + 1)  # 'Student ID'
                    rowData.append(self.studentFirstName[s])  # 'First Name'
                    rowData.append(self.studentLastName[s])  # 'Last Name'
                    rowData.append(self.studentFirstChoices[s])  # 'First choice'
                    rowData.append(self.studentSecondChoices[s])  # 'Second choice'
                    rowData.append(self.studentThirdChoices[s])  # 'Third choice'
                    rowData.append(ca)  # 'Course Assignment'
                    student_writer.writerow(rowData)
                    worksheet.write_row(row, 0, rowData)
                    row += 1
            worksheet.conditional_format('$G$2:$G$1048576', {'type': 'formula',
                                                'criteria': '=$G2=$D2',
                                                'format': highlight_green})
            worksheet.conditional_format('$G$2:$G$1048576', {'type': 'formula',
                                                'criteria': '=$G2=$E2',
                                                'format': highlight_yellow})
            worksheet.conditional_format('$G$2:$G$1048576', {'type': 'formula',
                                                'criteria': '=$G2=$F2',
                                                'format': highlight_red})
            worksheet.conditional_format('$D$2:$F$1048576', {'type': 'formula',
                                                'criteria': '=OR($D2=$E2,$D2=$F2,$E2=$F2,TYPE($D2)<>1,TYPE($E2)<>1,TYPE($F2)<>1)',
                                                'format': text_red})

        with open('Output_BulletVoting_Students.csv', mode='w', encoding='utf-8') as course_file:
            student_writer = writer(course_file, delimiter=',', quotechar='"', quoting=QUOTE_MINIMAL,
                                    lineterminator='\n')
            # write column names
            columnNames = []
            columnNames.append('Student ID')
            columnNames.append('First Name')
            columnNames.append('Last Name')
            columnNames.append('First choice')
            columnNames.append('Second choice')
            columnNames.append('Third choice')
            columnNames.append('Course Assignment')
            student_writer.writerow(columnNames)

            # write row data
            for s in range(self.S):
                ca = "%d" % self.studentIdAssignments[s][0]
                for i in range(1, len(self.studentIdAssignments[s])):
                    ca += ", %d" % self.studentIdAssignments[s][i]
                c1 = self.studentFirstChoices[s]
                c2 = self.studentSecondChoices[s]
                c3 = self.studentThirdChoices[s]

                if self.notUniqueChoices(c1, c2, c3):
                    rowData = []
                    rowData.append(s + 1)  # 'Student ID'
                    rowData.append(self.studentFirstName[s])  # 'First Name'
                    rowData.append(self.studentLastName[s])  # 'Last Name'
                    rowData.append(c1)  # 'First choice'
                    rowData.append(c2)  # 'Second choice'
                    rowData.append(c3)  # 'Third choice'
                    rowData.append(ca)  # 'Course Assignment'
                    student_writer.writerow(rowData)

        with open('Output_Unassigned_Students.csv', mode='w', encoding='utf-8') as course_file:
            student_writer = writer(course_file, delimiter=',', quotechar='"', quoting=QUOTE_MINIMAL,
                                    lineterminator='\n')
            worksheet = workbook.get_worksheet_by_name("Student Unassignments")
            # write column names
            columnNames = []
            columnNames.append('Student ID')
            columnNames.append('First Name')
            columnNames.append('Last Name')
            columnNames.append('First choice')
            columnNames.append('Second choice')
            columnNames.append('Third choice')
            student_writer.writerow(columnNames)
            worksheet.write_row('A1', columnNames)

            # write row data
            row = 1
            for s in range(self.S):
                if (self.studentIdAssignments[s][0] == -1):
                    rowData = []
                    rowData.append(s + 1)  # 'Student ID'
                    rowData.append(self.studentFirstName[s])  # 'First Name'
                    rowData.append(self.studentLastName[s])  # 'Last Name'
                    rowData.append(self.studentFirstChoices[s])  # 'First choice'
                    rowData.append(self.studentSecondChoices[s])  # 'Second choice'
                    rowData.append(self.studentThirdChoices[s])  # 'Third choice'
                    student_writer.writerow(rowData)
                    worksheet.write_row(row, 0, rowData)
                    row += 1
            worksheet.conditional_format('$D$2:$F$1048576', {'type': 'formula',
                                                'criteria': '=OR($D2=$E2,$D2=$F2,$E2=$F2,TYPE($D2)<>1,TYPE($E2)<>1,TYPE($F2)<>1)',
                                                'format': text_red})

        with open('Output_Stats.csv', mode='w', encoding='utf-8') as course_file:
            stats_writer = writer(course_file, delimiter=',', quotechar='"', quoting=QUOTE_MINIMAL, lineterminator='\n')
            worksheet = workbook.get_worksheet_by_name("Stats")

            # write column names
            rows = [totalWeightStat, firstChoiceStat, secondChoiceStat, thirdChoiceStat, noChoiceStat, multiChoiceStat, noAssignmentStat]
            for row in rows:
                stats_writer.writerow([row])
            worksheet.write_column('A1', rows)

        # done outputting
        workbook.close()


class HardConstraintMatcher(Matcher):
    def initVariables(self):
        super(HardConstraintMatcher, self).initVariables()
        self.classWillRun = [LpVariable("C%dr" % c, 0, 1, LpInteger)
                             for c in range(self.C)]

    def makeObjective(self):
        super(HardConstraintMatcher, self).makeObjective()
        objective = LpAffineExpression()
        for c in range(self.C):
            for s in range(self.S):
                objective += self.studentPreferences[s][c] * self.studentAssignments[s][c]
            # objective += (5*self.S/self.C)*self.classWillRun[c]

        # Add objective to model
        self.model += objective

    def makeConstraints(self):
        super(HardConstraintMatcher, self).makeConstraints()

        # Constraints described by Dr. Miller in email
        # runConstraints = [[LpConstraint() for s in range(self.S)] for c in range(self.C)]
        # for c in range(self.C):
        #     for s in range(self.S):  
        #         constraint = self.studentAssignments[s][c] - self.classWillRun[c]
        #         runConstraints[c][s] = LpConstraint(e=constraint, sense=-1, name="C%dS%dr"%(c, s), rhs=0)
        # sizeConstraints = [LpConstraint() for c in range(self.C)]    
        # for c in range(self.C):
        #     constraint = self.classWillRun[c] - self.sumStudentsInClass[c]
        #     sizeConstraints[c] = LpConstraint(e=constraint, sense=-1, name="C%ds"%c, rhs=0)

        # Constraints for each class
        classConstraints = [[LpConstraint() for c in range(self.C)] for i in range(2)]
        for c in range(self.C):
            hardMin = self.sumStudentsInClass[c] - self.courseMins[c] * self.classWillRun[c]
            hardMax = self.sumStudentsInClass[c] - self.courseMaxs[c] * self.classWillRun[c]
            minConstraint = LpConstraint(e=hardMin, sense=1, name="C%dm" % c, rhs=0)
            maxConstraint = LpConstraint(e=hardMax, sense=-1, name="C%dM" % c, rhs=0)
            classConstraints[0][c] = minConstraint
            classConstraints[1][c] = maxConstraint

        # Constraint limiting the number of classes a student should be assigned to
        maxAssignmentConstraint = [LpConstraint() for s in range(self.S)]
        for s in range(self.S):
            sumOfAssignments = LpAffineExpression()
            for c in range(self.C):
                sumOfAssignments += self.studentAssignments[s][c]
            maxAssignmentConstraint[s] = sumOfAssignments <= 1

        # Constraint ensuring that a student gets assigned to one of there three preferences or no class
        prefAssignmentConstraint = [LpConstraint() for s in range(self.S)]
        for s in range(self.S):
            sumOfNoPrefAssignments = LpAffineExpression()
            for c in range(self.C):
                if (self.studentPreferences[s][c] == 0):
                    sumOfNoPrefAssignments += self.studentAssignments[s][c]
            prefAssignmentConstraint[s] = sumOfNoPrefAssignments <= 0

        # Add constraints to model
        for s in range(self.S):
            self.model += maxAssignmentConstraint[s]
        for s in range(self.S):
            self.model += prefAssignmentConstraint[s]
        for c in range(self.C):
            # Dr. Miller's Constraints
            # for s in range(self.S):
            #     self.model += runConstraints[c][s]
            # self.model += sizeConstraints[c]
            for i in range(len(classConstraints)):
                self.model += classConstraints[i][c]
