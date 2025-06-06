
from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
from io import BytesIO
from datetime import timedelta
from difflib import SequenceMatcher

app = Flask(__name__)

def nomes_semelhantes(nome1, nome2):
    return SequenceMatcher(None, nome1.upper(), nome2.upper()).ratio() > 0.8

@app.route("/", methods=["GET"])
def index():
    return render_template_string(open("index.html").read())

@app.route("/upload", methods=["POST"])
def upload_files():
    notas = request.files["notas"]
    extrato = request.files["extrato"]

    try:
        df_notas = pd.read_excel(notas)
        df_extrato = pd.read_excel(extrato)

        # Tratamento básico
        df_notas["CNPJ"] = df_notas["CNPJ"].astype(str).str.replace(r'\D', '', regex=True)
        df_notas["Razão Social"] = df_notas["Razão Social"].astype(str).str.strip().str.upper()
        df_notas["Data de Vencimento"] = pd.to_datetime(df_notas["Data de Vencimento"])
        df_notas["Valor"] = df_notas["Valor"].astype(float)

        df_extrato["Data do Crédito"] = pd.to_datetime(df_extrato["Data do Crédito"])
        df_extrato["Valor"] = df_extrato["Valor"].astype(float)
        df_extrato["Descrição"] = df_extrato["Descrição"].astype(str).str.strip().str.upper()

        conciliado = []

        # Para cada lançamento no extrato
        for _, lanc in df_extrato.iterrows():
            data_lanc = lanc["Data do Crédito"]
            valor_lanc = lanc["Valor"]
            descricao = lanc["Descrição"]

            # Filtrar notas com data próxima e mesmo cliente
            candidatas = df_notas.copy()
            candidatas = candidatas[
                (abs(candidatas["Data de Vencimento"] - data_lanc) <= timedelta(days=5)) &
                (
                    (candidatas["CNPJ"].str[:8].isin([descricao.replace('.', '').replace('/', '').replace('-', '')[:8]])) |
                    (candidatas["Razão Social"].apply(lambda nome: nomes_semelhantes(nome, descricao)))
                )
            ]

            # Tentar encontrar combinações cujo somatório ≈ valor_lanc (tolerância até 10 reais)
            from itertools import combinations
            for i in range(1, len(candidatas) + 1):
                for combo in combinations(candidatas.index, i):
                    subset = candidatas.loc[list(combo)]
                    soma = subset["Valor"].sum()
                    if abs(soma - valor_lanc) <= 10:
                        for _, nf in subset.iterrows():
                            conciliado.append({
                                "Número NF": nf["Número NF"],
                                "CNPJ": nf["CNPJ"],
                                "Razão Social": nf["Razão Social"],
                                "Valor NF": nf["Valor"],
                                "Data Vencimento": nf["Data de Vencimento"].date(),
                                "Valor Crédito": valor_lanc,
                                "Data Crédito": data_lanc.date(),
                                "Descrição Crédito": descricao,
                                "Status": "Conciliado",
                                "Critério": "Match por CNPJ/Razão + Data + Soma ≈ Valor"
                            })
                        df_notas.drop(index=subset.index, inplace=True)
                        break
                else:
                    continue
                break

        # Adicionar NFs não conciliadas
        for _, nf in df_notas.iterrows():
            conciliado.append({
                "Número NF": nf["Número NF"],
                "CNPJ": nf["CNPJ"],
                "Razão Social": nf["Razão Social"],
                "Valor NF": nf["Valor"],
                "Data Vencimento": nf["Data de Vencimento"].date(),
                "Valor Crédito": "",
                "Data Crédito": "",
                "Descrição Crédito": "",
                "Status": "Não conciliado",
                "Critério": "Sem correspondência"
            })

        df_saida = pd.DataFrame(conciliado)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_saida.to_excel(writer, index=False, sheet_name="Conciliação")
        output.seek(0)

        return send_file(output, download_name="relatorio_conciliacao.xlsx", as_attachment=True)

    except Exception as e:
        return f"Erro ao processar arquivos: {e}"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
