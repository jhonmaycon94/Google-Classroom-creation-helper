import pyperclip, unicodedata

alunos_csv= pyperclip.paste().split("\n")

alunos = []

for aluno in alunos_csv:
    aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")
    alunos.append(aluno)
    
pyperclip.copy("\n".join(alunos))