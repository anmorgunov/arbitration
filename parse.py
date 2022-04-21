from json import load
from openpyxl import load_workbook

class Responses:

    def __init__(self, fname, code, jury):
        self.fname = fname
        self.CODE = code
        self.JURY = jury
        self.juryToStudents = {}
        self.parsed = []
        self.emailToApp = {}

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
                if (response:=ws[col+str(row)].value) is not None and response != '-':
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
                        self.juryToStudents[jury] = []
                    self.juryToStudents[jury].append(data)

    def _summary_for_jury(self):
        for jury, studs in self.juryToStudents.items():
            print(jury, len(studs))
                
    def main(self):
        self._parse_results()
        self._find_uniques() #oh god why
        self._make_queue_for_jury()
        self._summary_for_jury()



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