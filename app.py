
from flask import Flask, request, render_template_string
import pandas as pd
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template_string(open("index.html").read())

@app.route("/upload", methods=["POST"])
def upload_files():
    notas = request.files["notas"]
    extrato = request.files["extrato"]

    try:
        # Leitura dos arquivos Excel
        df_notas = pd.read_excel(notas)
        df_extrato = pd.read_excel(extrato)

        # Validação de colunas
        colunas_notas = ["CNPJ", "Razão Social", "Número NF", "Data de Vencimento", "Valor"]
        colunas_extrato = ["Data do Crédito", "Valor", "Descrição"]

        for col in colunas_notas:
            if col not in df_notas.columns:
                return f"Coluna obrigatória ausente em Notas: {col}"

        for col in colunas_extrato:
            if col not in df_extrato.columns:
                return f"Coluna obrigatória ausente em Extrato: {col}"

        # Tratamento dos dados
        df_notas["CNPJ"] = df_notas["CNPJ"].astype(str).str.replace(r'\D', '', regex=True)
        df_notas["Razão Social"] = df_notas["Razão Social"].astype(str).str.strip().str.upper()
        df_notas["Data de Vencimento"] = pd.to_datetime(df_notas["Data de Vencimento"])
        df_notas["Valor"] = df_notas["Valor"].astype(float)

        df_extrato["Data do Crédito"] = pd.to_datetime(df_extrato["Data do Crédito"])
        df_extrato["Valor"] = df_extrato["Valor"].astype(float)
        df_extrato["Descrição"] = df_extrato["Descrição"].astype(str).str.strip().str.upper()

        return (
            f"<h3>Arquivos tratados com sucesso!</h3>"
            f"<ul>"
            f"<li>Notas recebidas: {len(df_notas)}</li>"
            f"<li>Lançamentos no extrato: {len(df_extrato)}</li>"
            f"</ul>"
        )

    except Exception as e:
        return f"Erro ao processar arquivos: {e}"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
