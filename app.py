from flask import Flask, render_template, request, send_file
import pandas as pd
from itertools import combinations
import io
import re
import os

app = Flask(__name__)

def normalize_razao_social(razao):
    razao = str(razao)
    razao = re.sub(r'[^A-Z0-9 ]+', '', razao.upper())
    razao = re.sub(r'^\d+\s*', '', razao)
    return razao.strip()

def encontrar_combinacoes(notas, valor_alvo, tolerancia=3.0):
    resultados = []
    for r in range(1, len(notas) + 1):
        for combo in combinations(notas, r):
            soma = sum(n['Valor NF'] for n in combo)
            if abs(soma - valor_alvo) <= tolerancia:
                resultados.append(combo)
    return resultados

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    notas_file = request.files['notas']
    extrato_file = request.files['extrato']

    df_notas = pd.read_excel(notas_file)
    df_extrato = pd.read_excel(extrato_file)

    df_notas['Razao Normalizada'] = df_notas['Razão Social'].apply(normalize_razao_social)
    df_extrato['Razao Normalizada'] = df_extrato['Razão Social'].apply(normalize_razao_social)

    conciliados = []
    nao_conciliados = []

    for idx, pagamento in df_extrato.iterrows():
        valor_pago = pagamento['Valor']
        razao_pagador = pagamento['Razao Normalizada']

        possiveis_nfs = df_notas[df_notas['Razao Normalizada'].str.contains(razao_pagador, na=False)]
        lista_nfs = possiveis_nfs.to_dict('records')

        combinacoes = encontrar_combinacoes(lista_nfs, valor_pago)

        if combinacoes:
            melhor_combinacao = min(combinacoes, key=lambda c: abs(sum(n['Valor NF'] for n in c) - valor_pago))
            for nota in melhor_combinacao:
                conciliados.append({
                    **nota,
                    'Valor Pago': valor_pago,
                    'Data Pagamento': pagamento['Data'],
                    'Critério': 'Valor aproximado encontrado por combinação'
                })
        else:
            nao_conciliados.append({
                'Razão Social': pagamento['Razão Social'],
                'Valor Pago': valor_pago,
                'Data Pagamento': pagamento['Data'],
                'Status': 'Não conciliado'
            })

    df_conciliados = pd.DataFrame(conciliados)
    df_nao_conciliados = pd.DataFrame(nao_conciliados)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_conciliados.to_excel(writer, index=False, sheet_name='Conciliados')
        df_nao_conciliados.to_excel(writer, index=False, sheet_name='Nao Conciliados')
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="resultado_conciliacao.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8000))
    app.run(host='0.0.0.0', port=port)
