import pandas as pd
from tqdm import tqdm

clickup = pd.read_excel("BV ClickUp_Para Importação (1).xlsx", sheet_name='Tasks')
receb_pgt = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='BV')
liberacao = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='Liberacao de Retirada BV')

# Localiza
# clickup = pd.read_excel("Localiza ClickUp_Para Importação.xlsx", sheet_name='Tasks')
# receb_pgt = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='Localiza')

def atualizar_coluna(nome_coluna_destino, nome_coluna_fonte, df_fonte, df_destino):
    for index, row in clickup.iterrows():
        merged_df = pd.merge(df_destino, df_fonte[['Task Name', nome_coluna_fonte]], on='Task Name', how='left')
        # print(df_fonte['Task Name'])
        mask = (pd.isna(merged_df[nome_coluna_destino])) & (~pd.isna(merged_df[nome_coluna_fonte]))
        merged_df.loc[mask, nome_coluna_destino] = merged_df.loc[mask, nome_coluna_fonte]
        df_destino[nome_coluna_destino] = merged_df[nome_coluna_destino]

# Localiza

# atualizar_coluna('Veiculo - Modelo (naPista) (short text)', 'Veiculo - Modelo (naPista)', receb, clickup)
# atualizar_coluna('Taxa Administrativa - Valor (currency)', 'Taxa Administrativa - Valor', receb, clickup)
# atualizar_coluna('Comprador - Nome Razao Social (short text)', 'Comprador - Nome Razao Social', receb, clickup)
# atualizar_coluna('Comprador - CPF/CNPJ (short text)', 'Comprador - CPF/CNPJ', receb, clickup)
# atualizar_coluna('Boleto - Numero Controle  (short text)', 'Boleto - Numero de Controle', receb, clickup)
# atualizar_coluna('Taxa Administrativa - Data de Envio do Boleto (date)', 'Taxa Administrativa - Data de Envio do Boleto', receb, clickup)
# atualizar_coluna('Taxa Administrativa - Data de Pagamento do Boleto (date)', 'Taxa Administrativa - Data de Pagamento do Boleto', receb, clickup)
# atualizar_coluna('Taxa Administrativa - Data de Repasse (date)', 'Taxa Administrativa - Data de Repasse', receb, clickup)
# atualizar_coluna('Venda - Data da Venda (date)', 'Venda - Data da Venda', receb, clickup)
# atualizar_coluna('Taxa Administrativa - Data de Vencimento do Boleto  (date)', 'Taxa Administrativa - Data de Vencimento do Boleto', receb, clickup)

# BV

atualizar_coluna('Retirada - Pessoa Responsavel (short text)', 'Retirada - Pessoa Responsavel', liberacao, clickup)
atualizar_coluna('Retirada - CPF/CNPJ (short text)', 'Retirada - CPF', liberacao, clickup)
atualizar_coluna('Retirada - RG (short text)', 'Retirada - RG', liberacao, clickup)
atualizar_coluna('Repasse - Data de Repasse  (date)', 'Repasse - Data de Repasse', receb_pgt, clickup)
atualizar_coluna('Venda - Data de Pagamento (date)', 'Venda - Data de Pagamento', receb_pgt, clickup)
atualizar_coluna('ATPV-E - Data de Envio (date)', 'ATPV-E - Data de Envio', receb_pgt, clickup)
atualizar_coluna('ATPV-E - Data de Recebimento (date)', 'ATPV-E - Data de Recebimento', receb_pgt, clickup)

# clickup['ATPV-E - Data de Recebimento (date)'] = receb_pgt['ATPV-E - Data de Recebimento'].apply(lambda x: '1' if pd.notna(x) else x)
# clickup.rename(columns={'ATPV-E - Data de Recebimento': 'ATPV-E - Data de Recebimento (date)'}, inplace=True)

arqFinal = pd.DataFrame(clickup)
arqFinal.to_excel("base_final_BV_16_2.xlsx", index=False, sheet_name='BV')


