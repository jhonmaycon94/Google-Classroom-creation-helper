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

#database to compare

alunos_copiados = pyperclip.paste()
alunos_copiados = alunos_copiados.split("\n")
alunos_copiados = normatizar(alunos_copiados) 

final_list = []

qtd = 0

for aluno in alunos_copiados:
    match = [aluno_ws for aluno_ws in alunos_from_ws if aluno in aluno_ws]
    if not match:
        qtd = qtd + 1
        print(aluno)
print(str(qtd)+" ALUNOS\n")