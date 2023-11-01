import pandas as pd
import re
import datetime
import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


def buscaCaracterEspecial(inputString):
    pattern = r'[^a-zA-Z0-9\s]'
    return bool(re.search(pattern, inputString))


def removeCaracterEspecial(input_string):
    return re.sub(r'[^a-zA-Z0-9\s]', '', inputString)


def verificaArquivo(inputFile, arqOutput):
    try:
        workbook = openpyxl.load_workbook(inputFile)
        sheet = workbook.active

        for col in sheet.iter_cols(min_row=1, max_row=1):
            for cell in col:
                column_name = cell.value
                if column_name:
                    if buscaCaracterEspecial(column_name):
                        cleaned_column_name = removeCaracterEspecial(column_name)
                        cell.value = cleaned_column_name

        workbook.save(arqOutput)
        processarArquivo(arqOutput)
        messagebox.showinfo(title="Sucesso", message="Operação concluída com sucesso.")
    except Exception as e:
        print("Ocorreu um erro ao processar o arquivo XLSX:", str(e))
        messagebox.showerror(title="Erro", message="Ocorreu um erro ao processar o arquivo")

        
def processarArquivo(arqOutput):
    
    clickup = pd.read_excel(arquivo, sheet_name='Base ClickUp')
    receb_pgt = pd.read_excel(arquivo, sheet_name='Base Recebimento e Pagamento')
    liberacao = pd.read_excel(arquivo, sheet_name='Base Liberacao de Retirada')
    comissao = pd.read_excel(arquivo, sheet_name='Base de Comissao')
    
    placasAlteradas = []

    try:
        #liberações
        for index, row in clickup.iterrows():
            for idx, linha in liberacao.iterrows():
                if (row['Task Name'] == linha['Task Name '] and pd.isna(row['Retirada  Pessoa Responsvel short text']) == True):
                    arquivo.loc[index, 'Retirada  Pessoa Responsvel short text'] = liberacao.loc[idx, 'Retirada - Pessoa Responsavel']
                    placasAlteradas.append(row['Task Name'])

        for index, row in clickup.iterrows():
            for idx, linha in liberacao.iterrows():
                if (row['Task Name'] == linha['Task Name '] and pd.isna(row['Retirada  RG short text']) == True):
                    arquivo.loc[index, 'Retirada  RG short text'] = liberacao.loc[idx, 'Retirada - RG']
                    placasAlteradas.append(row['Task Name'])

        for index, row in clickup.iterrows():
            for idx, linha in comissao.iterrows():
                if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Taxa Administrativa  Valor currency']) == True):
                    arquivo.loc[index, 'Taxa Administrativa  Valor currency'] = comissao.loc[idx, 'Taxa Administrativa - Valor']
                    placasAlteradas.append(row['Task Name'])

        for index, row in clickup.iterrows():
            for idx, linha in comissao.iterrows():
                if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Comisso naPista currency']) == True):            
                    arquivo.loc[index, 'Comisso naPista currency'] = comissao.loc[idx, 'Comissao naPista']
                    placasAlteradas.append(row['Task Name'])

        for index, row in clickup.iterrows():
            for idx, linha in receb_pgt.iterrows():
                if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Comprador  Nome short text']) == True):
                    arquivo.loc[index, 'Comprador  Nome short text'] = receb_pgt.loc[idx, 'Comprador - Nome Razao Social']
                    placasAlteradas.append(row['Task Name'])
                    if (pd.isna(receb_pgt.loc[idx, 'Comprador - Nome Razao Social'] == False)):              
                        arquivo.loc[index, 'Venda  Confirmao drop down'] = "1"

        for index, row in clickup.iterrows():
            for idx, linha in receb_pgt.iterrows():
                if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Veculo  Chassi short text']) == True):                       
                    arquivo.loc[index, 'Veculo  Chassi short text'] = receb_pgt.loc[idx, 'Veiculo - Versao (naPista)']
                    placasAlteradas.append(row['Task Name'])

        for index, row in clickup.iterrows():
            for idx, linha in receb_pgt.iterrows():
                if (row['Task Name'] == linha['Task Name'] and pd.isna(row['ATPVE  Data de Envio date']) == True):                       
                    arquivo.loc[index, 'ATPVE  Data de Envio date'] = receb_pgt.loc[idx, 'ATPV-E - Data de Envio']       
                    placasAlteradas.append(row['Task Name'])    

        for index, row in clickup.iterrows():
            for idx, linha in receb_pgt.iterrows():
                if (row['Task Name'] == linha['Task Name'] and pd.isna(row['ATPVE  Data de Recebimento date']) == True):                              
                    arquivo.loc[index, 'ATPVE  Data de Recebimento date'] = receb_pgt.loc[idx, 'ATPV-E - Data de Recebimento']
                    placasAlteradas.append(row['Task Name'])
                    if (pd.isna(receb_pgt.loc[idx, 'ATPV-E - Data de Recebimento'] == False)):
                        arquivo.loc[index, 'ATPVE  Data de Recebimento date'] = "1"

        arqFinal = pd.DataFrame(clickup)
        arqFinal.to_excel("base_final.xlsx", index=False)

    except Exception as e:
            print("Ocorreu um erro ao processar o script", str(e))
            
def abrirArquivo():
    caminho_arquivo = filedialog.askopenfilename()
    if (caminho_arquivo != ""):
        caminho_output = "base_formatada_{datetime.datetime.now().strftime('%d_%H')}.xlsx"
        # caminho_label.config(text=f"Arquivo selecionado: {caminho_arquivo}")
        verificaArquivo(caminho_arquivo, caminho_output)
        return caminho_arquivo
        
# Inicialize a janela
root = tk.Tk()
root.title("Join - naPista")
root.iconbitmap("logo.4701e179iconNP.ico")
root.geometry("300x200")

# Crie um rótulo na janela
rotulo = tk.Label(root, text="Insira o arquivo principal:")
rotulo.pack()

# # Crie uma lista para exibir as informações
# lista_informacoes = tk.Listbox(root)
# lista_informacoes.pack()

# Crie um botão para abrir um arquivo
botao_abrir = tk.Button(root, text="Abrir Arquivo", command=abrirArquivo)
botao_abrir.pack()

# Crie um botão para processar o arquivo
botao_processar = tk.Button(root, text="Processar Arquivo", command=processarArquivo)
botao_processar.pack()

# Execute a janela principal
root.mainloop()