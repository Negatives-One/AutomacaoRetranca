import pandas
from datetime import datetime, timedelta
import tkinter as tk
import tkinter.filedialog as fd
from timeit import default_timer as timer

root = tk.Tk()
filez = fd.askopenfilenames(parent=root, title='Selecione as Planilhas')
if len(filez) < 1:
    exit()

savePath = fd.asksaveasfilename(parent=root, title='Onde Salvar?', filetypes = [('XLSX', '.xlsx')], defaultextension='.xlsx')
if savePath == "":
    exit()

start = timer()

colunas = ["PROGRAMA", "DATA", "TÍTULO DA OBRA MUSICAL",
           "TIPO DE MÚSICA", "AUTOR", "INTERPRETE", "SEGUNDOS", "CLASSIFICAÇÃO"]

base = pandas.DataFrame(columns=colunas)

dataFrames = []

print("Lendo Arquivos...")
for i in range(len(filez)):
    print("Leitura:", int(i / len(filez) * 100), "%")
    dataFrames.append(pandas.read_excel(filez[i], header=2,sheet_name=2))

def DeleteUselessColumns(table):
    toDelete = []
    # table['Product_Name'].values[i]
    for i in range(2, 9):
        toDelete.append(i)
    for j in range(15, 20):
        toDelete.append(j)
    table = table.drop(table.columns[toDelete], axis=1)
    return table

def DeleteUselessRows(table):
    toDrop = []
    for i in range(len(table.index)):
        value = table.iat[i, 0]
        if type(value) == float:
            toDrop.append(i)
        if type(value) == str:
            if "BLOCO" in value:
                toDrop.append(i)
    table = table.drop(toDrop)
    return table

for n in range(len(dataFrames)):
    print("Formatação:", int(n / len(dataFrames) * 100), "%")
    dataFrames[n] = DeleteUselessColumns(dataFrames[n])
    dataFrames[n] = DeleteUselessRows(dataFrames[n])
    dataFrames[n] = dataFrames[n].reset_index(drop=True)

def AddDataToMain():
    print("Consolidando...")
    totalIndex = 0
    placed = 0
    for i in range(len(dataFrames)):
        print("Calculando Espaço: ", int(i / len(dataFrames) * 100), "%")
        totalIndex += len(dataFrames[i].index)
    for k in range(totalIndex):
        print("Alocando Espaço: ", int(k / totalIndex * 100), "%")
        base.loc[base.shape[0]] = [None] * len(colunas)
    for j in range(len(dataFrames)):
        print("Adicionando Valores: ", int(j / len(dataFrames) * 100), "%")
        for CCount in range(len(colunas)):
            for LCount in range(len(dataFrames[j].index)):
                base.iloc[placed + LCount,
                          CCount] = dataFrames[j].iloc[LCount, CCount]
        placed += len(dataFrames[j].index)

def SerialNumberToDateTime(coluna):
    dataInicial = '1900/01/01'
    for i in range(len(coluna)):
        date = datetime.strptime(dataInicial, "%Y/%m/%d")
        modified_date = date + timedelta(days=int(coluna[i]))
        coluna[i] = str(datetime.strftime(modified_date, "%d/%m/%Y"))
    return coluna

def AdjustAutors(coluna):
    for i in range(len(coluna)):
        texto = str(coluna[i])
        texto = texto.replace(',', ' /')
        coluna[i] = texto
    return coluna
    

AddDataToMain()

base = base.sort_values(by =["PROGRAMA", "DATA"] )
base = base.reset_index(drop=True)   
base["DATA"] = SerialNumberToDateTime(base["DATA"])
base["AUTOR"] = AdjustAutors(base["AUTOR"])

try:
    writer = pandas.ExcelWriter(savePath)

    base.to_excel(writer, 'Dataset 1')

    writer.save()

    end = timer()
    print("Done!\n" + str(int(end - start)) + " segundos de execução")
    exit()
finally:
    exit()