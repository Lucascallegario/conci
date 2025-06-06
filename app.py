
from flask import Flask, request, render_template_string
import pandas as pd

app = Flask(__name__)

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
        return f"Arquivos recebidos com sucesso!<br>Notas: {len(df_notas)} registros<br>Extrato: {len(df_extrato)} registros"
    except Exception as e:
        return f"Erro ao processar arquivos: {e}"

if __name__ == "__main__":
    app.run(debug=True)
