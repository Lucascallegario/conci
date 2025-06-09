import os
import unicodedata
import pandas as pd
import jellyfish
from flask import Flask, request, render_template_string, send_file
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from itertools import combinations

app = Flask(__name__)
app.static_folder = 'static'
HISTORICO_DIR = "historico"
os.makedirs(HISTORICO_DIR, exist_ok=True)

def nomes_parecidos(a, b, threshold=0.8):
    return SequenceMatcher(None, a.upper(), b.upper()).ratio() >= threshold

def normalizar(texto):
    texto = texto.upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = texto.replace("PAGAMENTO", "").replace("DEPÓSITO", "").replace("DEPOSITO", "")
    return " ".join(texto.split())

def extrair_raiz(nome):
    texto = normalizar(nome)
    palavras = texto.split()
    ignorar = {"LTDA", "EIRELI", "ME", "MAT", "MATRIZ", "FILIAL", "COMERCIAL", "NACIONAL", "CENTRO", "SHOPPING", "SUL", "NORTE", "UNIDADE", "LOJA", "GRUPO"}
    return " ".join([p for p in palavras if p not in ignorar])

@app.route("/", methods=["GET"])
def index():
    arquivos = sorted(os.listdir(HISTORICO_DIR), reverse=True)
    return render_template_string(open("index.html").read(), arquivos=arquivos)

@app.route("/upload", methods=["POST"])
def upload_files():
    notas = request.files["notas"]
    extrato = request.files["extrato"]
    try:
        df_notas = pd.read_excel(notas)
        df_extrato = pd.read_excel(extrato)

        df_notas["Raiz"] = df_notas["Razão Social"].apply(extrair_raiz)
        df_notas["Data de Vencimento"] = pd.to_datetime(df_notas["Data de Vencimento"])
        df_notas["Valor"] = df_notas["Valor"].astype(float)

        df_extrato["Data do Crédito"] = pd.to_datetime(df_extrato["Data do Crédito"])
        df_extrato["Valor"] = df_extrato["Valor"].astype(float)
        df_extrato["Descrição"] = df_extrato["Descrição"].astype(str).str.upper()
        df_extrato["Raiz"] = df_extrato["Descrição"].apply(extrair_raiz)

        conciliado = []
        df_pendentes = df_notas.copy()

        for _, lanc in df_extrato.iterrows():
            data_lanc = lanc["Data do Crédito"]
            valor_lanc = lanc["Valor"]
            raiz_lanc = lanc["Raiz"]

            candidatas = df_pendentes[
                (abs(df_pendentes["Data de Vencimento"] - data_lanc) <= timedelta(days=5)) &
                (df_pendentes["Raiz"] == raiz_lanc)
            ]

            match = False
            for i in range(1, len(candidatas)+1):
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
                                "Descrição Crédito": lanc["Descrição"],
                                "Status": "Conciliado",
                                "Critério": "Agrupamento por raiz"
                            })
                        df_pendentes.drop(index=subset.index, inplace=True)
                        match = True
                        break
                if match: break

        for _, lanc in df_extrato.iterrows():
            if df_pendentes.empty:
                break
            matches = []
            for _, nf in df_pendentes.iterrows():
                score = 0
                if nf["Raiz"] == lanc["Raiz"]:
                    score += 5
                if abs(nf["Valor"] - lanc["Valor"]) <= 10:
                    score += 4
                if abs((nf["Data de Vencimento"] - lanc["Data do Crédito"]).days) <= 5:
                    score += 3
                jaro_sim = jellyfish.jaro_winkler_similarity(str(nf["Razão Social"]), str(lanc["Descrição"]))
                if jaro_sim >= 0.90:
                    score += 2
                if score >= 10:
                    matches.append((score, nf))
            if matches:
                best_match = sorted(matches, key=lambda x: -x[0])[0][1]
                conciliado.append({
                    "Número NF": best_match["Número NF"],
                    "CNPJ": best_match["CNPJ"],
                    "Razão Social": best_match["Razão Social"],
                    "Valor NF": best_match["Valor"],
                    "Data Vencimento": best_match["Data de Vencimento"].date(),
                    "Valor Crédito": lanc["Valor"],
                    "Data Crédito": lanc["Data do Crédito"].date(),
                    "Descrição Crédito": lanc["Descrição"],
                    "Status": "Conciliado",
                    "Critério": "Pontuação Probabilística"
                })
                df_pendentes.drop(index=best_match.name, inplace=True)

        for _, nf in df_pendentes.iterrows():
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
        conciliados_data_credito = set((c["Data Crédito"], c["Valor Crédito"]) for c in conciliado if c["Status"] == "Conciliado")
        extrato_nao_conciliado = df_extrato[~df_extrato.apply(lambda row: (row["Data do Crédito"].date(), row["Valor"]) in conciliados_data_credito, axis=1)].copy()

        filename = f"conciliacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(HISTORICO_DIR, filename)

        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            # Aba 1: Conciliação
            df_saida.to_excel(writer, index=False, sheet_name="Conciliação")
            ws = writer.sheets["Conciliação"]
            ws.autofilter(0, 0, len(df_saida), len(df_saida.columns) - 1)
            ws.freeze_panes(1, 0)

            status_col = df_saida.columns.get_loc("Status")
            for i, status in enumerate(df_saida["Status"], start=1):
                color = "#D9EAD3" if status == "Conciliado" else "#F4CCCC"
                ws.set_row(i, None, writer.book.add_format({"bg_color": color}))

            # Aba 2: Extrato não conciliado
            extrato_nao_conciliado.to_excel(writer, index=False, sheet_name="Extrato não conciliado")
            ws2 = writer.sheets["Extrato não conciliado"]
            ws2.autofilter(0, 0, len(extrato_nao_conciliado), len(extrato_nao_conciliado.columns) - 1)
            ws2.freeze_panes(1, 0)

            # Aba 3: Resumo
            total_nfs = len(df_saida)
            n_conciliadas = sum(df_saida["Status"] == "Conciliado")
            n_nao = total_nfs - n_conciliadas
            perc = round(n_conciliadas / total_nfs * 100, 2) if total_nfs else 0
            valor_conc = df_saida[df_saida["Status"] == "Conciliado"]["Valor Crédito"].sum()
            valor_nao = df_saida[df_saida["Status"] != "Conciliado"]["Valor NF"].sum()

            df_resumo = pd.DataFrame({
                "Métrica": [
                    "Total de NFs", "Conciliadas", "Não conciliadas",
                    "% Conciliação", "Valor conciliado", "Valor não conciliado"
                ],
                "Valor": [
                    total_nfs, n_conciliadas, n_nao, f"{perc}%", valor_conc, valor_nao
                ]
            })
            df_resumo.to_excel(writer, index=False, sheet_name="Resumo")
            ws3 = writer.sheets["Resumo"]
            ws3.set_column("A:B", 25)

        return send_file(filepath, download_name=filename, as_attachment=True)
    except Exception as e:
        return f"Erro ao processar: {e}"

@app.route("/historico/<nome>")
def baixar_arquivo(nome):
    return send_file(os.path.join(HISTORICO_DIR, nome), as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
