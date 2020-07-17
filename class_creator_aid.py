import pyperclip, re, unicodedata, openpyxl

def normatizar(alunos):
    alunos_formatados = []
    for aluno in alunos:
        aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")
        aluno = aluno.upper().strip()
        alunos_formatados.append(aluno)
    return alunos_formatados

def extract_students_from_copy(copy):
    alunos = re.findall(r"(\b[A-Z][^\d]{1,}\t\t)", copy)
    alunos = normatizar(alunos)
    if not alunos:
        copy = str(copy).split("\n")
        copy = normatizar(copy)
        copy = "\n".join(copy)
        alunos = re.findall(r"((?<=\d\t)[A-Z ]{3,})", copy)
    return alunos  

def load_students_from_sheet(sheet):
    alunos = []
    for i in range(1, sheet.max_row):
        alunos.append(sheet.cell(row=i, column=1).value)
    return alunos

def search_student_in_sheet(copied_students, alunos_from_sheet):
    alunos = []
    for aluno in copied_students:
        match = [aluno_ws for aluno_ws in alunos_from_sheet if aluno in aluno_ws]
        if not match:
            continue
        alunos.append(match[0])
    return alunos

def alunos_without_email(alunos):
    sem_email = []
    for aluno in alunos:
        if "@" not in aluno:
            sem_email.append(aluno)
    return sem_email

def alunos_to_clipboard(alunos):
    pyperclip.copy("\n".join(alunos))

def format_output(alunos, alunos_sem_email):
    qnt = len(alunos)
    print(str(qnt)+" Alunos copiados para área de transferência\n")
    if alunos_sem_email: 
        print("ALUNOS SEM EMAIL:\n")
        for aluno in alunos_sem_email:
            print(aluno)

wb = openpyxl.load_workbook("emails academicos.xlsx")
ws_alunos = wb['Planilha1']
alunos_from_ws = load_students_from_sheet(ws_alunos)
alunos_from_ws = normatizar(alunos_from_ws)    

copied_students = pyperclip.paste()
alunos = extract_students_from_copy(copied_students)

alunos_para_copiar = search_student_in_sheet(alunos, alunos_from_ws)
alunos_sem_email = alunos_without_email(alunos_para_copiar)
alunos_to_clipboard(alunos_para_copiar)
format_output(alunos_para_copiar, alunos_sem_email)



