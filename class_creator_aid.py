import pyperclip, re, unicodedata

tb_alunos = pyperclip.paste()

alunos = re.findall(r"(\b[A-Z][^\d]{1,}\t\t)", tb_alunos)

for aluno in alunos:
    aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")
    pyperclip.copy(aluno)
    input("pressione enter para o próximo nome")



