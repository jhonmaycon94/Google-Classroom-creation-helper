import pyperclip, re, unicodedata, time
from pynput.keyboard import Listener

tb_alunos = pyperclip.paste()

alunos = re.findall(r"(\b[A-Z][^\d]{1,}\t\t)", tb_alunos)

for aluno in alunos:
    aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")

i = 0

def on_press(key):
    global i
    try:
        if key == 'p':
            pyperclip.copy(alunos[i].strip())
            if i < len(alunos):
                i=i+1
    except AttributeError:
        pass     

def on_release(key):
    pass
with Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()

