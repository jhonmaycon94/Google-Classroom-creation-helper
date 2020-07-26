import unicodedata
import re
import pyperclip
import openpyxl


class GoogleClassCreationHelper:

    @staticmethod
    def normatize(students): 
        for indice, student in enumerate(students):
            # tira acentos e ç dos nomes dos alunos
            student = unicodedata.normalize(
                "NFKD", student.strip()).encode(
                    "ascii", "ignore").decode("ascii")
            student = student.upper().strip()
            students[indice] = student
        return students

    def extract_students_from_SIGAA(self, clipboard_copy):
        # Regex para extrair todos alunos do integrado/subsequente
        students = re.findall(r"(\b[A-Z][^\d]{1,}\t\t)", clipboard_copy)
        students = self.normatize(students)

        '''
        Se não conseguir extrair alunos da string,
        checa se os alunos são de engenharia
        usando uma Regex diferente
        '''
        if not students:
            clipboard_copy = self.normatize(str(clipboard_copy).split("\n"))
            clipboard_copy = "\n".join(clipboard_copy)
            students = re.findall(r"((?<=\d\t)[A-Z ]{3,})", clipboard_copy)
        return students

    @staticmethod
    def load_students_from_sheet():
        workbook = openpyxl.load_workbook("emails academicos.xlsx")
        sheet = workbook['Planilha1']
        students = []
        last_row = sheet.max_row
        for i in range(1, last_row):
            student = sheet.cell(row=i, column=1).value
            students.append(student)
        return students

    @staticmethod
    def search_student_in_sheet(students_from_copy, students_from_sheet):
        students = [] 
        for student in students_from_copy:
            students.extend([student_ws for student_ws in students_from_sheet if student in student_ws])
        return students

    @staticmethod
    def students_without_email(students):
        no_email = []
        for student in students:
            if "@" not in student:
                no_email.append(student)
        return no_email

    @staticmethod
    def students_to_clipboard(students):
        pyperclip.copy("\n".join(students))

    def execute(self):
        students_from_sheet = self.normatize(self.load_students_from_sheet())
        students = self.extract_students_from_SIGAA(pyperclip.paste())

        self.students = self.search_student_in_sheet(students, students_from_sheet)
        self.students_without_email = self.students_without_email(self.students)
        self.students_to_clipboard(self.students)
