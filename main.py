import tkinter as tk
from tkinter import ttk
import openpyxl


#BACKEND
def load_data():
    path = r"C:\Users\setto\OneDrive\Ambiente de Trabalho\cadastro\funcionarios.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)

    for col_name in cols:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[0:]:
        treeview.insert('', tk.END, values=value_tuple)


def insert_data():
    name = name_entry.get()
    age = int(age_spinbox.get())
    gender = status_gender.get()
    nationality = "Estrangeiro" if a.get() else "Nacional"
    
    if name == '' or name == 'Nome':
        print("Nome é obrigatorio")
    else:
        #Salvar dados na folha excel
        path = r"C:\Users\setto\OneDrive\Ambiente de Trabalho\cadastro\funcionarios.xlsx"  
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        row_values = [name, age, gender, nationality]
        sheet.append(row_values)
        workbook.save(path)

        #Inserir dado salvo na treeview
        treeview.insert('', tk.END, values=row_values)

        #Limpar e colocar os dados iniciais dos nossos campos
        name_entry.delete(0, 'end')
        name_entry.insert(0, "Nome")

        age_spinbox.delete(0, 'end')
        age_spinbox.insert(0, "Idade")

        status_gender.set(gender_list[0])

        checkbutton.state(["!selected"])
    

def mode_theme():
    if mode_switch.instate(["selected"]):
        style.theme_use('forest-light')
    else:
        style.theme_use('forest-dark')



#FRONTEND
#iniciando a nossa janela
root = tk.Tk()
root.title('Cadastro de Funcionarios')


#configuracao do nosso tema
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

#frame principal
frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insira os dados")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

#Criando os campos do nosso formulario de cadastro.
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Nome")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry. grid(row=0, column=0, padx=5, pady=(30, 5), sticky='ew')

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=70)
age_spinbox.insert(0, "Idade")
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

gender_list = ["Masculino", "Feminino", "Outro"]
status_gender = ttk.Combobox(widgets_frame, values=gender_list)
status_gender.current(0)
status_gender. grid(row=2, column=0, padx=5, pady=5, sticky='ew')

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Estrangeiro", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky='nsew')

button = ttk.Button(widgets_frame, text='Inserir dados'.upper(), command=insert_data)
button.grid(row=4, column=0, padx=5, pady=5, sticky='nsew')

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=5, pady=10, sticky='ew')

mode_switch = ttk.Checkbutton(widgets_frame, text='Modo', style='Switch', command=mode_theme)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky='nsew')

#Criando a nossa treeview para dados do Banco
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, padx=(0, 20), pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Nome", "Idade", "Genero", "Nacionalidade")
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=13)

treeview.column("Nome", width=120)
treeview.column("Idade", width=30)
treeview.column("Genero", width=100)
treeview.column("Nacionalidade", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)

load_data()


#Rodando a nossa janela
root.mainloop()
