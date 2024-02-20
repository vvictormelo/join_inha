import pandas as pd
import datetime
import openpyxl
import sys
import time
from tqdm import tqdm

def processaExcel():
    
    clickup = pd.read_excel("BV ClickUp_Para Importação (1).xlsx", sheet_name='Tasks')
    receb_pgt = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='BV')
    liberacao = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='Liberacao de Retirada BV')
    comissao = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='BV')

    placasAlteradas = []

    try:
        for index, row in tqdm(clickup.iterrows(), total=len(clickup), desc='Processando'):
            for index, row in clickup.iterrows():
                for idx, linha in liberacao.iterrows():
                    if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Retirada - Pessoa Responsavel (short text)']) == True):
                        clickup.loc[index, 'Retirada - Pessoa Responsavel (short text)'] = liberacao.loc[idx, 'Retirada - Pessoa Responsavel']
                        placasAlteradas.append(row['Task Name'])

            # for index, row in clickup.iterrows():
            #     for idx, linha in liberacao.iterrows():
            #         if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Retirada - CPF/CNPJ (short text)']) == True):
            #             clickup.loc[index, 'Retirada - CPF/CNPJ (short text)'] = liberacao.loc[idx, 'Retirada - CPF']
            #             placasAlteradas.append(row['Task Name'])
                        
            # for index, row in clickup.iterrows():
            #     for idx, linha in liberacao.iterrows():
            #         if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Retirada - RG (short text)']) == True):
            #             clickup.loc[index, 'Retirada - RG (short text)'] = liberacao.loc[idx, 'Retirada - RG']
            #             placasAlteradas.append(row['Task Name'])

            # for index, row in clickup.iterrows():
            #     for idx, linha in comissao.iterrows():
            #         if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Repasse - Data de Repasse  (date)']) == True):            
            #             clickup.loc[index, 'Repasse - Data de Repasse  (date)'] = comissao.loc[idx, 'Repasse - Data de Repasse']
            #             placasAlteradas.append(row['Task Name'])
                        
            # for index, row in clickup.iterrows():
            #     for idx, linha in comissao.iterrows():
            #         if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Venda - Data de Pagamento (date)']) == True):            
            #             clickup.loc[index, 'Venda - Data de Pagamento (date)'] = comissao.loc[idx, 'Venda - Data de Pagamento']
            #             placasAlteradas.append(row['Task Name'])

            # for index, row in clickup.iterrows():
            #     for idx, linha in receb_pgt.iterrows():
            #         if (row['Task Name'] == linha['Task Name'] and pd.isna(row['ATPV-E - Data de Envio (date)']) == True):                       
            #             clickup.loc[index, 'ATPV-E - Data de Envio (date)'] = receb_pgt.loc[idx, 'ATPV-E - Data de Envio']       
            #             placasAlteradas.append(row['Task Name'])

            # for index, row in clickup.iterrows():
            #     for idx, linha in receb_pgt.iterrows():
            #         if (row['Task Name'] == linha['Task Name'] and pd.isna(row['ATPV-E - Data de Recebimento (date)']) == True):                              
            #             clickup.loc[index, 'ATPV-E - Data de Recebimento (date)'] = receb_pgt.loc[idx, 'ATPV-E - Data de Recebimento']
            #             placasAlteradas.append(row['Task Name'])
            #             if (pd.isna(receb_pgt.loc[idx, 'ATPV-E - Data de Recebimento'] == False)):
            #                 clickup.loc[index, 'ATPV-E - Data de Recebimento (date)'] = "1"

        arqFinal = pd.DataFrame(clickup)
        arqFinal.to_excel("base_final.xlsx", index=False)
        print(placasAlteradas)
        
        time.sleep(0.1)

    except Exception as e:
            print("Ocorreu um erro ao processar o script", str(e))

if __name__ == "__main__":
    processaExcel()