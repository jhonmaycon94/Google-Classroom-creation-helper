import pyperclip, re, unicodedata, time, openpyxl

wb = openpyxl.load_workbook("emails academicos.xlsx")
ws_alunos = wb['Planilha1']
alunos_from_ws = []

for i in range(1, ws_alunos.max_row):
    alunos_from_ws.append(ws_alunos.cell(row=i, column=1).value)

def normatizar(alunos):
    alunos_formatados = []
    for aluno in alunos:
        aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")
        aluno = aluno.upper().strip()
        alunos_formatados.append(aluno)

    return alunos_formatados

alunos_from_ws = normatizar(alunos_from_ws)

tb_alunos = pyperclip.paste()

alunos_unicode = re.findall(r"(\b[A-Z][^\d]{1,}\t\t)", tb_alunos)
if not alunos_unicode:
    tb_alunos = str(tb_alunos).split("\n")
    tb_alunos = normatizar(tb_alunos)
    tb_alunos = "\n".join(tb_alunos)
    alunos_unicode = re.findall(r"((?<=\d\t)[A-Z ]{3,})", tb_alunos)

alunos = normatizar(alunos_unicode) 

final_list = []
alunos_sem_email = []

qtd = 0

for aluno in alunos:
    match = [aluno_ws for aluno_ws in alunos_from_ws if aluno in aluno_ws]
    if not match:
        continue 
    qtd = qtd + 1
    match = "\n".join(match)
    final_list.append(str(qtd)+" "+match)
    if '@' not in match:
        alunos_sem_email.append(match)
    
pyperclip.copy("\n".join(final_list))

print(str(qtd)+" Alunos copiados para área de transferencia\n")
if alunos_sem_email:
    print("ALUNOS SEM EMAIL:\n")
    for aluno in alunos_sem_email:
        print(aluno)


