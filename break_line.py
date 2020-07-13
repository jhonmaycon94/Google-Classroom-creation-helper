import pyperclip

alunos = pyperclip.paste()

alunos = alunos.replace(",","\n")

pyperclip.copy(alunos)

