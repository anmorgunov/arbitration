from multiprocessing.spawn import get_preparation_data
from numpy import gradient
from openpyxl import load_workbook
from openpyxl import Workbook
import helper

class Responses:

    def __init__(self, fname, code, jury):
        self.fname = fname
        self.CODE = code
        self.JURY = jury
        self.juryToStudents = {}
        self.parsed = []
        self.emailToApp = {}
        self.juryToGrToData = {}

        # screw it, hardcoding is king
        self.PROBCOLS = 'GHIJKLMNOP'

    def _parse_results(self):
        wb = load_workbook(self.fname)
        ws = wb['sheet']
        row = 2
        while True:
            if ws['A'+str(row)].value is None:
                break
            student = {}
            for col, param in self.CODE.items():
                if (response:=ws[col+str(row)].value) is not None and response != '-' and response != '.' and response != 'Нет вопросов':
                    student[param] = response
            self.parsed.append(student)
            row += 1

    def _find_uniques(self):
        for student in self.parsed:
            if (email:=student['email']) in self.emailToApp:
                print('someone has applied more than once')
            if email not in self.emailToApp:
                self.emailToApp[email] = {}
            for param, data in student.items():
                if param != 'email':
                    if param not in self.emailToApp[email]:
                        self.emailToApp[email][param] = []
                    self.emailToApp[email][param].append(data)

    def _make_queue_for_jury(self):
        for email, data in self.emailToApp.items():
            grade = data['grade'][0]
            for probCol in self.PROBCOLS:
                if (prob:=self.CODE[probCol]) in data:
                    studResponse = data[prob]
                    # print(grade, email, probCol)
                    jury = self.JURY[grade][prob]
                    # print(jury)
                    if jury not in self.juryToStudents:
                        self.juryToStudents[jury] = {}
                    if email not in self.juryToStudents[jury]:
                        self.juryToStudents[jury][email] = {}
                    for param, values in data.items():
                        if param not in self.juryToStudents[jury][email]:
                            self.juryToStudents[jury][email][param] = []
                        self.juryToStudents[jury][email][param].append(values)
                    # self.juryToStudents[jury].append(data)

    def _summary_for_jury(self):
        juryToGrToNum = {}
        for jury, emailToParam in self.juryToStudents.items():
            print(jury, len(emailToParam))
            if jury not in juryToGrToNum:
                juryToGrToNum[jury] = {}
            for email, paramToData in emailToParam.items():
                grade = paramToData['grade'][0][0] #hardcoding but screw it
                if grade not in juryToGrToNum[jury]:
                    juryToGrToNum[jury][grade] = []
                juryToGrToNum[jury][grade].append(paramToData)
        
        for jury, grToData in juryToGrToNum.items():
            for gr, data in grToData.items():
                print(jury, gr, len(data))
        
        self.juryToGrToData = juryToGrToNum

    def _create_the_queue(self):
        wb = Workbook()
        for grade in (9, 10, 11):
            wb.create_sheet(f"{grade} класс")
            ws = wb[f"{grade} класс"]
            col = 'A'
            for jury, grToData in self.juryToGrToData.items():
                row = 1
                ws[col + str(row)] = jury
                row += 1
                if grade in grToData:
                    for student in grToData[grade]:
                        # print(student['name'])
                        ws[col + str(row)] = row - 1 
                        ws[helper.getNextCol(col) + str(row)] = student['name'][0][0] 
                        row += 1 
                col = helper.getNextCol(col)
                col = helper.getNextCol(col)
        wb.save('queue.xlsx')

    def _is_there_a_conflict(self, data):
        for grade, juryToStudents in data.items():
            for jury, students in juryToStudents.items():
                for i, student in enumerate(students):
                    for jury2, students2 in juryToStudents.items():
                        if jury2 != jury:
                            if len(students2) > 0 and i < len(students2):
                                if students2[i] == student:
                                    return True
        return False

    def _remove_repetitions_in_queue(self):
        wb = load_workbook('queue.xlsx')
        grToData = {}
        for grade in (9, 10, 11):
            ws = wb[f"{grade} класс"]
            juryToStuds = {}
            col = 'A'
            while True:
                if (jury := ws[col+str(1)].value) is None:
                    break

                row = 1
                
                if jury not in juryToStuds:
                    juryToStuds[jury] = []
                    row = 2
                    while True:
                        if (name:=ws[helper.getNextCol(col)+str(row)].value) is None:
                            break
                        juryToStuds[jury].append(name)
                        row += 1
                grToData[grade] = juryToStuds
                col = helper.getNextCol(col)
                col = helper.getNextCol(col)
        
        while True:
            conflict = self._is_there_a_conflict(grToData)
            print(conflict)
            if not conflict:
                break
            for grade, juryToStudents in grToData.items():
                for jury, students in juryToStudents.items():
                    for i, student in enumerate(students):
                        for jury2, students2 in juryToStudents.items():
                            if jury2 != jury:
                                if len(students2) > 0 and i < len(students2):
                                    if students2[i] == student:
                                        students = students[:i] + students[i+1:] + students[i:i+1]
                                        grToData[grade][jury] = students
                                        break
        new_wb = Workbook()
        for grade, juryToStudents in grToData.items():
            new_wb.create_sheet(f"{grade} класс")
            ws = new_wb[f"{grade} класс"]
            col = 'A'
            for jury, students in juryToStudents.items():
                # print(students)
                row = 1
                ws[col + str(row)] = jury
                row += 1
                for student in students:
                    # print(student['name'])
                    ws[col + str(row)] = row - 1 
                    ws[helper.getNextCol(col) + str(row)] = student
                    row += 1 
                col = helper.getNextCol(col)
                col = helper.getNextCol(col)
        new_wb.save('queue_with_no_conflicts.xlsx')






                
    def main(self):
        # self._parse_results()
        # self._find_uniques() #oh god why
        # self._make_queue_for_jury()
        # self._summary_for_jury()
        # self._create_the_queue()
        self._remove_repetitions_in_queue()


COL_TO_PARAM = {
    'B': 'email',
    'C': 'oblast',
    'E': 'name',
    'F': 'grade',
    'G': 'p1',
    'H': 'p2',
    'I': 'p3',
    'J': 'p4',
    'K': 'p5',
    'L': 'p6',
    'M': 'p7',
    'N': 'p8',
    'O': 'p2-1',
    'P': 'p2-2', # GHIJKLMNOP
}

P_TO_JURY = {
    9: {
        'p1': 'Тасанов А.',
        'p2': 'Тасанов А.',
        'p3': 'Черданцев В.',
        'p4': 'Моргунов А.',
        'p5': 'Тайшыбай А.',
        'p6': 'Черданцев В.',
        'p7': 'Молдагулов Г.',
        'p2-1': 'Молдагулов Г.'
    },
    10: {
        'p1': 'Бекхожин Ж.',
        'p2': 'Моргунов А.',
        'p3': 'Тайшыбай А.',
        'p4': 'Загрибельный Б.',
        'p5': 'Мадиева М.',
        'p6': 'Молдагулов Г.',
        'p7': 'Бекхожин Ж.',
        'p2-1': 'Бекхожин Ж.',
        'p2-2': 'Черданцев В.'
    },
    11: {
        'p1': 'Мадиева М.',
        'p2': 'Моргунов А.',
        'p3': 'Тайшыбай А.',
        'p4': 'Загрибельный Б.',
        'p5': 'Мадиева М.',
        'p6': 'Моргунов А.',
        'p7': 'Тайшыбай А.',
        'p8': 'Загрибельный Б.',
        'p2-1': 'Молдагулов Г.',
        'p2-2': 'Черданцев В.'
    }
}

# assume sheet is names "sheet"
chObj = Responses('chemistry.xlsx', COL_TO_PARAM, P_TO_JURY)
chObj.main()