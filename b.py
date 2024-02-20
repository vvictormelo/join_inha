import pandas as pd

# Leitura dos arquivos
clickup = pd.read_excel("BV ClickUp_Para Importação (1).xlsx", sheet_name='Tasks')
receb_pgt = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='BV')
liberacao = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='Liberacao de Retirada BV')
comissao = pd.read_excel("Planilha Backoffice_Dados que constam.xlsx", sheet_name='BV')

placasAlteradas = []

# Atualiza as colunas com base nas condições
def update_column(target_df, source_df, target_column, source_column):
    mask = target_df[target_column].isna()
    merge_df = target_df.merge(source_df[['Task Name', source_column]], on='Task Name', suffixes=('', '_src'), how='left')
    target_df[target_column] = merge_df[target_column].fillna(merge_df[f'{source_column}_src'])
    placasAlteradas.extend(merge_df.loc[mask, 'Task Name'])

# Atualização das colunas específicas usando itertuples
def update_columns_using_itertuples(target_df, source_df, target_column, source_column):
    mask = target_df[target_column].isna()
    source_dict = {row['Task Name']: row[source_column] for row in source_df.itertuples(index=False)}
    target_df[target_column] = target_df.apply(lambda row: source_dict.get(row['Task Name'], row[target_column]), axis=1)
    placasAlteradas.extend(target_df.loc[mask, 'Task Name'])

# Atualização das colunas específicas usando itertuples
update_columns_using_itertuples(clickup, liberacao, 'Retirada - Pessoa Responsavel (short text)', 'Retirada - Pessoa Responsavel')
update_columns_using_itertuples(clickup, liberacao, 'Retirada - CPF/CNPJ (short text)', 'Retirada - CPF')
update_columns_using_itertuples(clickup, liberacao, 'Retirada - RG (short text)', 'Retirada - RG')
update_columns_using_itertuples(clickup, comissao, 'Repasse - Data de Repasse  (date)', 'Repasse - Data de Repasse')
update_columns_using_itertuples(clickup, comissao, 'Venda - Data de Pagamento (date)', 'Venda - Data de Pagamento')
update_columns_using_itertuples(clickup, receb_pgt, 'ATPV-E - Data de Envio (date)', 'ATPV-E - Data de Envio')
update_columns_using_itertuples(clickup, receb_pgt, 'ATPV-E - Data de Recebimento (date)', 'ATPV-E - Data de Recebimento')

# Define 'ATPV-E - Data de Recebimento (date)' como '1' se não for NaN no DataFrame de origem
clickup['ATPV-E - Data de Recebimento (date)'] = clickup['ATPV-E - Data de Recebimento'].apply(lambda x: '1' if pd.notna(x) else x)

# Cria um DataFrame final e salva como um arquivo Excel
arqFinal = pd.DataFrame(clickup)
arqFinal.to_excel("base_final.xlsx", index=False)
