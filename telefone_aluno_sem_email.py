import pyperclip, re, unicodedata, time, openpyxl, csv

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

#planilha consulta geral
cg_file = open('consulta_geral.csv')
cg_file_output = open('lista_alunos.csv', 'w', newline='')
output_writer = csv.writer(cg_file_output)
cg_reader = csv.reader(cg_file)
cg_dados_alunos = list(cg_reader)

for lista in cg_dados_alunos:
    lista = normatizar(lista)

for i in range(1, ws_alunos.max_row):
    alunos_from_ws.append(ws_alunos.cell(row=i, column=1).value)

alunos_from_ws = normatizar(alunos_from_ws)

for aluno in alunos_from_ws:
    if '@' not in aluno:
        for lista in cg_dados_alunos:
            if aluno in str(lista):
                output_writer.writerow(lista)
                break
cg_file_output.close()                