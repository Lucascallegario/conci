
from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
from io import BytesIO
from datetime import timedelta
from difflib import SequenceMatcher
from openai_conciliador import gerar_prompt_ia, consultar_ia

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

        df_notas["CNPJ"] = df_notas["CNPJ"].astype(str).str.replace(r'\D', '', regex=True)
        df_notas["Razão Social"] = df_notas["Razão Social"].astype(str).str.strip().str.upper()
        df_notas["Data de Vencimento"] = pd.to_datetime(df_notas["Data de Vencimento"])
        df_notas["Valor"] = df_notas["Valor"].astype(float)

        df_extrato["Data do Crédito"] = pd.to_datetime(df_extrato["Data do Crédito"])
        df_extrato["Valor"] = df_extrato["Valor"].astype(float)
        df_extrato["Descrição"] = df_extrato["Descrição"].astype(str).str.strip().str.upper()

        conciliado = []
        df_pendentes = df_notas.copy()

        for _, lanc in df_extrato.iterrows():
            data_lanc = lanc["Data do Crédito"]
            valor_lanc = lanc["Valor"]
            descricao = lanc["Descrição"]

            candidatas = df_pendentes[
                (abs(df_pendentes["Data de Vencimento"] - data_lanc) <= timedelta(days=5)) &
                (
                    (df_pendentes["CNPJ"].str[:8].isin([descricao.replace('.', '').replace('/', '').replace('-', '')[:8]])) |
                    (df_pendentes["Razão Social"].apply(lambda nome: nomes_semelhantes(nome, descricao)))
                )
            ]

            from itertools import combinations
            match_encontrado = False

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
                                "Critério": "Match Regras Fixas"
                            })
                        df_pendentes.drop(index=subset.index, inplace=True)
                        match_encontrado = True
                        break
                if match_encontrado:
                    break

            if not match_encontrado and len(df_pendentes) > 0:
                candidatas_ia = df_pendentes.head(5)
                notas = []
                for _, row in candidatas_ia.iterrows():
                    notas.append({
                        "Numero": row["Número NF"],
                        "Razao": row["Razão Social"],
                        "Valor": row["Valor"],
                        "Vencimento": row["Data de Vencimento"]
                    })

                prompt = gerar_prompt_ia(descricao, valor_lanc, data_lanc, notas)
                resposta = consultar_ia(prompt)
                sugestao = []
                for nota in notas:
                    if str(nota["Numero"]) in resposta:
                        sugestao.append(nota["Numero"])

                for _, nf in candidatas_ia.iterrows():
                    status = "Conciliado por IA" if nf["Número NF"] in sugestao else "Não conciliado"
                    conciliado.append({
                        "Número NF": nf["Número NF"],
                        "CNPJ": nf["CNPJ"],
                        "Razão Social": nf["Razão Social"],
                        "Valor NF": nf["Valor"],
                        "Data Vencimento": nf["Data de Vencimento"].date(),
                        "Valor Crédito": valor_lanc if status == "Conciliado por IA" else "",
                        "Data Crédito": data_lanc.date() if status == "Conciliado por IA" else "",
                        "Descrição Crédito": descricao if status == "Conciliado por IA" else "",
                        "Status": status,
                        "Critério": "IA – agrupamento sem match fixo" if status == "Conciliado por IA" else "Sem correspondência"
                    })
                    if status == "Conciliado por IA":
                        df_pendentes.drop(index=nf.name, inplace=True)

        df_saida = pd.DataFrame(conciliado)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_saida.to_excel(writer, index=False, sheet_name="Conciliação")
            workbook = writer.book
            worksheet = writer.sheets["Conciliação"]

            # Formatação de colunas
            format_nf = workbook.add_format({'bg_color': '#F2F2F2'})
            format_banco = workbook.add_format({'bg_color': '#DAE8FC'})
            format_status = workbook.add_format({'bg_color': '#D9EAD3'})

            worksheet.set_column("A:E", 18, format_nf)
            worksheet.set_column("F:H", 18, format_banco)
            worksheet.set_column("I:J", 25, format_status)

            # Adicionar autofiltro
            worksheet.autofilter(0, 0, len(df_saida), len(df_saida.columns) - 1)

        output.seek(0)
        return send_file(output, download_name="relatorio_conciliacao.xlsx", as_attachment=True)

    except Exception as e:
        return f"Erro ao processar arquivos: {e}"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
