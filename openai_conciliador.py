
import openai
import os

openai.api_key = os.environ.get("OPENAI_API_KEY")

def gerar_prompt_ia(descricao_credito, valor_credito, data_credito, notas):
    lista_formatada = ""
    for i, nf in enumerate(notas, 1):
        lista_formatada += f"{i}. NF {nf['Numero']} – {nf['Razao']} – R$ {nf['Valor']:.2f} – Venc. {nf['Vencimento'].strftime('%Y-%m-%d')}\n"

    prompt = f'''
Você é um especialista em análise de pagamentos empresariais e conciliação bancária com foco em redes varejistas com múltiplas filiais.

Receba a descrição de um lançamento bancário e uma lista de notas fiscais pendentes de um grupo econômico. Com base nas informações, identifique quais notas fiscais foram provavelmente quitadas por esse lançamento.

O objetivo é sugerir agrupamentos de notas com base em:
- Nomes parecidos ou que remetam à mesma rede (ex: Supermercado ABC Matriz, Supermercado ABC Nacional, Supermercado ABC Rio).
- Soma de valores próxima ao valor recebido (até R$10,00 de diferença).
- Datas de vencimento próximas da data do crédito (até 5 dias antes ou depois).

Sempre considere que a matriz pode pagar pelas filiais.

### Descrição do crédito bancário:
"{descricao_credito}"

### Valor do crédito recebido:
R$ {valor_credito:.2f}

### Data do crédito:
{data_credito.strftime('%Y-%m-%d')}

### Notas fiscais pendentes:

{lista_formatada}

Para cada nota fiscal sugerida, responda:
- Se ela parece pertencer ao grupo que pagou
- Por que você acha isso (ex: nome semelhante, valor, proximidade de data, padrão de rede)

No final, responda com uma lista dos números das NFs que você acredita terem sido quitadas nesse recebimento.

Formato de resposta:
NF: [número], Justificativa: [explicação curta]
(... repetir ...)
Sugeridas: [número1, número2, número3]
'''.strip()

    return prompt


def consultar_ia(prompt):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=500
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        return f"Erro na consulta à OpenAI: {e}"
