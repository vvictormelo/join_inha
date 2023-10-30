from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

planilha = 'base importação.xlsx'

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("files")
        clickup = pd.read_excel(planilha, sheet_name='Base ClickUp')
        receb_pgt = pd.read_excel(planilha, sheet_name='Base Recebimento e Pagamento')
        liberacao = pd.read_excel(planilha, sheet_name='Base Liberacao de Retirada')
        comissao = pd.read_excel(planilha, sheet_name='Base de Comissao')
        
        missing_info = []

        for file in files:
            df = pd.read_excel(file)

            #liberações
            for index, row in clickup.iterrows():
                for idx, linha in liberacao.iterrows():
                    if (row['Task Name'] == linha['Task Name '] and pd.isna(row['Retirada - Pessoa Responsavel (short text)']) == True):
                        clickup.loc[index, 'Retirada - Pessoa Responsavel (short text)'] = liberacao.loc[idx, 'Retirada - Pessoa Responsavel']
                        clickup.loc[index, 'Retirada - RG (short text)'] = liberacao.loc[idx, 'Retirada - RG']

            #Comissões
            for index, row in clickup.iterrows():
                for idx, linha in comissao.iterrows():
                    if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Taxa Administrativa - Valor (currency)']) == True):
                        clickup.loc[index, 'Taxa Administrativa - Valor (currency)'] = comissao.loc[idx, 'Taxa Administrativa - Valor']
                        clickup.loc[index, 'Comissao naPista (currency)'] = comissao.loc[idx, 'Comissao naPista']

            #Pagamentos
            for index, row in clickup.iterrows():
                for idx, linha in receb_pgt.iterrows():
                    if (row['Task Name'] == linha['Task Name'] and pd.isna(row['Comprador - Nome (short text)']) == True):
                        
                        clickup.loc[index, 'Veiculo - Versao (naPista) (short text)'] = receb_pgt.loc[idx, 'Veiculo - Versao (naPista)']
                        clickup.loc[index, 'ATPV-E - Data de Envio (date)'] = receb_pgt.loc[idx, 'ATPV-E - Data de Envio']
                        clickup.loc[index, 'ATPV-E - Data de Recebimento (date)'] = receb_pgt.loc[idx, 'ATPV-E - Data de Recebimento']
                        clickup.loc[index, 'Comprador - Nome (short text)'] = receb_pgt.loc[idx, 'Comprador - Nome Razao Social']
                        
                        if (pd.isna(receb_pgt.loc[idx, 'ATPV-E - Data de Recebimento'] == False)):
                            clickup.loc[index, 'Envio ATPV-E - Confirmacao (drop down)'] = "1"
                            
                        if (pd.isna(receb_pgt.loc[idx, 'Comprador - Nome Razao Social'] == False)):              
                            clickup.loc[index, 'Venda - Confirmacao (drop down)'] = "1"

        if missing_info:
            missing_info_text = "\n".join(missing_info)
            return render_template("result.html", missing_info=missing_info_text)
        else:
            return render_template("result.html", missing_info="Nenhum campo ausente encontrado.")

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
