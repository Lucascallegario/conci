
import unicodedata
import itertools
from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
from difflib import SequenceMatcher

app = Flask(__name__)
app.static_folder = 'static'
HISTORICO_DIR = "historico"
os.makedirs(HISTORICO_DIR, exist_ok=True)

def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = texto.replace("PAGAMENTO", "").replace("DEPÓSITO", "").replace("DEPOSITO", "")
    return " ".join(texto.split())

def nomes_semelhantes(nome1, nome2, threshold=0.8):
    nome1 = normalizar(nome1)
    nome2 = normalizar(nome2)
    return (nome1 in nome2) or (nome2 in nome1) or SequenceMatcher(None, nome1, nome2).ratio() >= threshold

@app.route("/", methods=["GET"])
def index():
    arquivos = sorted(os.listdir(HISTORICO_DIR), reverse=True)
    return render_template_string("""
        <!DOCTYPE html>
        <html lang="pt-br">
        <head>
            <meta charset="UTF-8">
            <title>Conciliação by Lucas</title>
            <style>
                body { font-family: Arial, sans-serif; background: #f8f9fa; margin: 0; padding: 0; text-align: center; }
                .container { max-width: 700px; margin: 50px auto; background: #ffffff; padding: 40px; box-shadow: 0 0 10px rgba(0,0,0,0.05); border-radius: 12px; }
                h1 { color: #333; margin-bottom: 30px; }
                label { font-weight: bold; display: block; margin-top: 20px; margin-bottom: 8px; }
                input[type="file"] { padding: 6px; }
                input[type="submit"] { margin-top: 30px; padding: 10px 25px; font-size: 16px; background-color: #0066cc; color: white; border: none; border-radius: 6px; cursor: pointer; }
                input[type="submit"]:hover { background-color: #004999; }
                .historico { margin-top: 40px; }
                a { color: #0066cc; text-decoration: none; }
                a:hover { text-decoration: underline; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Conciliação by Lucas</h1>
                <form action="/upload" method="post" enctype="multipart/form-data">
                    <label>Notas Fiscais (Contas a Receber):</label>
                    <input type="file" name="notas" required>
                    <label>Extrato Bancário:</label>
                    <input type="file" name="extrato" required>
                    <br>
                    <input type="submit" value="Iniciar Conciliação">
                </form>
                <div class="historico">
                    <h2>Histórico de Conciliações</h2>
                    {% for arquivo in arquivos %}
                        <p><a href="/historico/{{ arquivo }}">{{ arquivo }}</a></p>
                    {% endfor %}
                </div>
            </div>
        </body>
        </html>
    """, arquivos=arquivos)

@app.route("/upload", methods=["POST"])
def upload_files():
    notas_file = request.files["notas"]
    extrato_file = request.files["extrato"]
    df_notas = pd.read_excel(notas_file)
    df_extrato = pd.read_excel(extrato_file)

    df_notas["Razao Normalizada"] = df_notas["Razão Social"].apply(normalizar)
    df_extrato["Descricao Normalizada"] = df_extrato["Descrição"].apply(normalizar)
    df_notas["Valor"] = df_notas["Valor"].astype(float)
    df_extrato["Valor"] = df_extrato["Valor"].astype(float)

    conciliado = []
    usadas = set()

    for idx_banco, row_banco in df_extrato.iterrows():
        if idx_banco in usadas:
            continue
        valor_banco = row_banco["Valor"]
        candidatos = df_notas[~df_notas.index.isin(usadas)]

        for r in range(1, len(candidatos)+1):
            for combo in itertools.combinations(candidatos.index, r):
                subset = candidatos.loc[list(combo)]
                soma = subset["Valor"].sum()
                if abs(soma - valor_banco) <= 3:
                    if all(nomes_semelhantes(row_banco["Descricao Normalizada"], x) for x in subset["Razao Normalizada"]):
                        for _, row_nf in subset.iterrows():
                            conciliado.append({
                                "Número NF": row_nf["Número NF"],
                                "CNPJ": row_nf["CNPJ"],
                                "Razão Social": row_nf["Razão Social"],
                                "Valor NF": row_nf["Valor"],
                                "Valor Crédito": valor_banco,
                                "Data Crédito": row_banco["Data do Crédito"],
                                "Descrição Crédito": row_banco["Descrição"],
                                "Status": "Conciliado",
                                "Critério": "Soma de múltiplas NFs"
                            })
                        usadas.update(combo)
                        usadas.add(idx_banco)
                        break
            if idx_banco in usadas:
                break

    for idx_nf, row_nf in df_notas.iterrows():
        if idx_nf not in usadas:
            conciliado.append({
                "Número NF": row_nf["Número NF"],
                "CNPJ": row_nf["CNPJ"],
                "Razão Social": row_nf["Razão Social"],
                "Valor NF": row_nf["Valor"],
                "Valor Crédito": "",
                "Data Crédito": "",
                "Descrição Crédito": "",
                "Status": "Não conciliado",
                "Critério": "Sem correspondência"
            })

    df_resultado = pd.DataFrame(conciliado)
    filename = f"conciliacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = os.path.join(HISTORICO_DIR, filename)
    df_resultado.to_excel(filepath, index=False)
    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route("/historico/<nome>")
def baixar_arquivo(nome):
    return send_file(os.path.join(HISTORICO_DIR, nome), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
