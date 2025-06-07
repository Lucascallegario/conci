
from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
from io import BytesIO
from datetime import datetime, timedelta
from difflib import SequenceMatcher

app = Flask(__name__)
app.static_folder = 'static'
HISTORICO_DIR = "historico"
os.makedirs(HISTORICO_DIR, exist_ok=True)

def nomes_semelhantes(nome1, nome2):
    return SequenceMatcher(None, nome1.upper(), nome2.upper()).ratio() > 0.8


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

    palavras = nome.upper().split()
    ignorar = {"LTDA", "EIRELI", "ME", "MAT", "MATRIZ", "FILIAL", "COMERCIAL", "NACIONAL", "CENTRO", "SHOPPING", "SUL", "NORTE"}
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

            from itertools import combinations
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
        filename = f"conciliacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(HISTORICO_DIR, filename)

        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            df_saida.to_excel(writer, index=False, sheet_name="Conciliação")
            ws = writer.sheets["Conciliação"]
            ws.autofilter(0, 0, len(df_saida), len(df_saida.columns) - 1)
            ws.set_column("A:E", 18)
            ws.set_column("F:H", 18)
            ws.set_column("I:J", 25)

        return send_file(filepath, download_name=filename, as_attachment=True)
    except Exception as e:
        return f"Erro: {e}"

@app.route("/historico/<nome>")
def baixar_arquivo(nome):
    return send_file(os.path.join(HISTORICO_DIR, nome), as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
