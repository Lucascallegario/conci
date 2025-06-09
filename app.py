import os
import itertools
import pandas as pd
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'resultados'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def normalizar_razao(razao):
    if not isinstance(razao, str):
        return ""
    return ' '.join([palavra for palavra in razao.upper().split() if not palavra.isnumeric()])

def buscar_combinacoes(possibilidades, valor_alvo, tolerancia=3):
    for r in range(1, len(possibilidades) + 1):
        for combo in itertools.combinations(possibilidades, r):
            soma = sum(combo)
            if abs(soma - valor_alvo) <= tolerancia:
                return combo
    return None

@app.route('/')
def index():
    arquivos = os.listdir(RESULT_FOLDER)
    return render_template('index.html', arquivos=arquivos)

@app.route('/upload', methods=['POST'])
def upload():
    file_notas = request.files['notas']
    file_extrato = request.files['extrato']
    notas_path = os.path.join(UPLOAD_FOLDER, secure_filename(file_notas.filename))
    extrato_path = os.path.join(UPLOAD_FOLDER, secure_filename(file_extrato.filename))
    file_notas.save(notas_path)
    file_extrato.save(extrato_path)

    df_notas = pd.read_excel(notas_path)
    df_extrato = pd.read_excel(extrato_path)

    df_notas['Razao Normalizada'] = df_notas['Razão Social'].apply(normalizar_razao)
    df_extrato['Razao Normalizada'] = df_extrato['Razão Social'].apply(normalizar_razao)

    df_notas['Conciliado'] = False
    df_notas['Grupo'] = None

    resultados = []
    for cliente in df_extrato['Razao Normalizada'].unique():
        recebimentos_cliente = df_extrato[df_extrato['Razao Normalizada'] == cliente]
        notas_cliente = df_notas[(df_notas['Razao Normalizada'] == cliente) & (~df_notas['Conciliado'])]

        valores_nf = list(notas_cliente['Valor NF'])
        indices_nf = list(notas_cliente.index)

        for _, linha in recebimentos_cliente.iterrows():
            valor_pago = linha['Valor Crédito']
            combinacao = buscar_combinacoes(valores_nf, valor_pago)
            if combinacao:
                usados = 0
                for val in combinacao:
                    for idx in indices_nf:
                        if not df_notas.at[idx, 'Conciliado'] and abs(df_notas.at[idx, 'Valor NF'] - val) < 0.01:
                            df_notas.at[idx, 'Conciliado'] = True
                            df_notas.at[idx, 'Grupo'] = f"{cliente} | {valor_pago:.2f} em {linha['Data Crédito'].date()}"
                            usados += 1
                            break
                if usados != len(combinacao):
                    continue

    df_notas['Status'] = df_notas['Conciliado'].apply(lambda x: 'Conciliado' if x else 'Não conciliado')
    nome_saida = f"resultado_conciliacao_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    caminho_saida = os.path.join(RESULT_FOLDER, nome_saida)
    df_notas.to_excel(caminho_saida, index=False)

    return send_file(caminho_saida, as_attachment=True)

@app.route('/historico/<arquivo>')
def historico(arquivo):
    caminho = os.path.join(RESULT_FOLDER, arquivo)
    return send_file(caminho, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
