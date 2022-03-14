import os
import pandas
import tkinter as tk
import tkinter.filedialog as fd

exceto = ["ARQUIVO"]

def SelectFolder():
    folderMonth = fd.askdirectory(parent=root, title="Selecione o mês.")
    if len(folderMonth) < 1:
        exit()

root = tk.Tk()
root.title("Automação Planilha Arquivo")
root.geometry("500x300")
buttonSelectFolder = tk.Button(root, text="Selecionar pasta do mês", command=SelectFolder)
buttonSelectFolder.pack()
buttonSelectFolder.place(x=160, y=40)
buttonSelectFolder.bg
root.mainloop()
# folderMonth = fd.askdirectory(parent=root, title="Selecione o mês.")
# if len(folderMonth) < 1:
#     exit()

# def get_immediate_subdirectories(a_dir):
#     return [name for name in os.listdir(a_dir) if os.path.isdir(os.path.join(a_dir, name))]

# days = get_immediate_subdirectories(folderMonth)

# savePath = fd.asksaveasfilename(parent=root, title='Onde Salvar?', filetypes = [('XLSX', '.xlsx')], defaultextension='.xlsx')
# if savePath == "":
#     exit()

# colunas = ["RETRANCA", "DIA", "REPORTER"]

# base = pandas.DataFrame(columns=colunas)

# base = base.sort_values(by =["DIA", "REPORTER"] )

# def AddDataToMain():
#     #FOR DOS DIAS
#     for i in range(len(days)):
#         reporters = get_immediate_subdirectories(folderMonth + "/" + days[i])
#         if reporters.__contains__("ARQUIVO"):
#             reporters.remove("ARQUIVO")
#         new_row = {'RETRANCA':" ", 'DIA': " ", 'REPORTER':" "}
#         base.loc[len(base.index)] = new_row
#         #FOR DOS REPORTERS
#         for j in range(len(reporters)):
#             retrancas = get_immediate_subdirectories(folderMonth + "/" + days[i] + "/" + reporters[j])
#             #FOR DAS RETRANCAS
#             for n in range(len(retrancas)):
#                 new_row = {'RETRANCA':retrancas[n], 'DIA': days[i], 'REPORTER':reporters[j]}
#                 base.loc[len(base.index)] = new_row

# AddDataToMain()

# writer = pandas.ExcelWriter(savePath)

# base.to_excel(writer, 'Dataset 1')

# writer.save()