import tkinter as tk
import tkinter.scrolledtext as scrolledtext
import creatorAid
import openpyxl
import pyperclip

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        #Title
        self.master.title("Classroom Creator Aid")

        #Frame
        self.mainFrame = tk.Frame(self.master, bg='white')
        self.mainFrame.pack_propagate(0)
        self.mainFrame.pack(fill=tk.BOTH, expand=1)

        #Label
        self.output_lb = tk.Label(self.mainFrame, 
                                  text="Lista de alunos copiados para área de transferencia:",
                                  font = ('Swis721 Blk BT',9),
                                  anchor = tk.W,
                                  bg='#0F59A8',
                                  fg='white',
                                  padx=5
                                  )
        self.output_lb.pack(anchor=tk.SW, pady=(10, 0), padx=20)

        #TextArea
        self.output_txt = scrolledtext.ScrolledText(self.mainFrame,
                                                    width=60, 
                                                    height=16,
                                                    font=('Nirmala UI',8),
                                                    pady=10,
                                                    padx=10,
                                                    bg='#CCDBDB',
                                                    borderwidth=1,
                                                    relief='sunken')                                                  
        self.output_txt.configure(state='disabled')
        self.output_txt.pack(padx=20, fill=tk.X)

        #Label
        self.lb_sem_email = tk.Label(self.mainFrame, 
                                     text='Alunos sem email : ',
                                     font=('Swis721 Blk BT',10),
                                     bg='red',
                                     fg='white'
                                     )
        self.lb_sem_email.pack(anchor=tk.W, padx=20, pady=(50, 0))
        
        #TextArea
        self.output_sem_email = scrolledtext.ScrolledText(self.mainFrame,
                                                          width=50,
                                                          pady=10,
                                                          padx=10,
                                                          font= ('Nirmala UI',8),
                                                          height=10,
                                                          bg='#CCDBDB',
                                                          borderwidth = 1,
                                                          relief='sunken')
        self.output_sem_email.configure(state='disabled')
        self.output_sem_email.pack(side=tk.LEFT, pady=(0, 20), padx=(20, 0))

        #Button
        self.magic_button = tk.Button(self.mainFrame, 
                                      text="RUN",
                                      bg='#CCDBDB', 
                                      command=self.process, 
                                      font=('Swis721 Blk BT',10),
                                      )
        self.magic_button.pack(side=tk.RIGHT, padx=100, anchor=tk.S, pady=20)

    def process(self):
        #Limpa TextArea
        self.output_txt.configure(state='normal')
        self.output_txt.delete('1.0', 'end')

        self.output_sem_email.configure(state='normal')
        self.output_sem_email.delete('1.0', 'end')

        creatorAid.execute()

        #Alunos TextArea
        qtd=0
        for aluno in creatorAid.alunos:
            qtd=qtd+1
            self.output_txt.insert('end', str(qtd)+".  "+aluno+"\n\n")
        self.output_txt.configure(state='disabled')   
        self.output_txt.update()

        #Alunos sem email TextArea
        qtd=0
        for aluno in creatorAid.alunos_s_email:
            qtd=qtd+1
            self.output_sem_email.insert('end', str(qtd)+".  "+aluno+"\n\n")
        self.output_sem_email.configure(state='disabled')
        self.output_sem_email.update()


    
root = tk.Tk()
root.geometry('620x500')
app = Application(master=root)
app.mainloop()