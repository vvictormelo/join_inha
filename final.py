import pandas as pd

# Leitura dos arquivos Excel
clickup = pd.read_excel("BV ClickUp_Para Importação (1).xlsx", sheet_name='Tasks')
receb_pgt = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='BV')
liberacao = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='Liberacao de Retirada BV')

# Função para atualizar uma coluna no DataFrame de destino com base em valores de outra fonte de dados
def atualizar_coluna(nome_coluna_destino, nome_coluna_fonte, df_fonte, df_destino):
    # Mescla os DataFrames com base no nome da tarefa ('Task Name')
    merged_df = pd.merge(df_destino, df_fonte[['Task Name', nome_coluna_fonte]], on='Task Name', how='left')
    # Cria uma máscara para identificar as células na coluna destino que estão vazias e têm um valor correspondente na fonte de dados
    mask = (pd.isna(merged_df[nome_coluna_destino])) & (~pd.isna(merged_df[nome_coluna_fonte]))
    # Atualiza as células na coluna destino com os valores correspondentes da fonte de dados
    df_destino.loc[mask, nome_coluna_destino] = merged_df.loc[mask, nome_coluna_fonte]


# Atualiza as colunas no DataFrame clickup com base nos valores dos DataFrames receb_pgt e liberacao
atualizar_coluna('Retirada - Pessoa Responsavel (short text)', 'Retirada - Pessoa Responsavel', liberacao, clickup)
atualizar_coluna('Retirada - CPF/CNPJ (short text)', 'Retirada - CPF', liberacao, clickup)
atualizar_coluna('Retirada - RG (short text)', 'Retirada - RG', liberacao, clickup)
atualizar_coluna('Repasse - Data de Repasse  (date)', 'Repasse - Data de Repasse', receb_pgt, clickup)
atualizar_coluna('Venda - Data de Pagamento (date)', 'Venda - Data de Pagamento', receb_pgt, clickup)
atualizar_coluna('ATPV-E - Data de Envio (date)', 'ATPV-E - Data de Envio', receb_pgt, clickup)
atualizar_coluna('ATPV-E - Data de Recebimento (date)', 'ATPV-E - Data de Recebimento', receb_pgt, clickup)


# Salva o DataFrame atualizado em um novo arquivo Excel
clickup.to_excel("base_final_BV_16_2.xlsx", index=False, sheet_name='BV')
