import pyperclip, re, unicodedata, time

op = ""

while op != 's':
    tb_alunos = pyperclip.paste()

    alunos = re.findall(r"(\b[A-Z][^\d]{1,}\t\t)", tb_alunos)
    
    qnt = 0
    for aluno in alunos:
        aluno = unicodedata.normalize("NFKD", aluno.strip()).encode("ascii", "ignore").decode("ascii")
        pyperclip.copy(aluno)
        qnt = qnt + 1
        print("na área de transferencia: " + aluno +"\n")
        input("tecle enter para o próximo nome...\n")
        print(str(qnt) + " aluno(s) copiado(s) \n")
        
    print("Turma Finalizada: "+ str(qnt) + " copiados\n")    
    op = input("tecle s para sair ou enter para colar nova turma\n")    

