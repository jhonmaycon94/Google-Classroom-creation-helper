import pyperclip, re, unicodedata, time, openpyxl

def normatizar(alunos):
    alunos_formatados = []
    for aluno in alunos:
        aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")
        aluno = aluno.upper().strip()
        alunos_formatados.append(aluno)
    return alunos_formatados

#original database
wb = openpyxl.load_workbook("emails academicos.xlsx")
ws_alunos = wb['Planilha1']
alunos_from_ws = []

for i in range(1, ws_alunos.max_row):
    alunos_from_ws.append(ws_alunos.cell(row=i, column=1).value)

alunos_from_ws = normatizar(alunos_from_ws)

for aluno in alunos_from_ws:
    if '@' not in aluno:
        print(aluno)