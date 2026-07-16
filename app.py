import os
from flask import Flask, render_template
import pandas as pd

# 1. Descobre o caminho da pasta onde o app.py está localizado
app = Flask(__name__, template_folder='Templates', static_folder='Static')

# Caminho seguro para o Excel na pasta raiz
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
caminho_excel = os.path.join(diretorio_atual, "Controle Financeiro 2026.xlsx")

def formata_real(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return f"R$ {valor}"

@app.route('/')
def home():
    try:
        # Verifica se o arquivo do Excel realmente subiu junto com o código
        if not os.path.exists(caminho_excel):
            return f"<h1>Erro: Arquivo Excel não encontrado</h1><p>O arquivo 'Controle Financeiro 2026.xlsx' não foi detectado no servidor. Certifique-se de que ele está na mesma pasta do seu app.py.</p>"

        # Lendo o Excel com o caminho seguro
        df = pd.read_excel(caminho_excel, sheet_name="Dados", skiprows=9)
        
        # Limpeza e organização dos dados (conforme o formato da sua planilha)
        if df.columns[0].startswith('Unnamed'):
            df = df.drop(columns=[df.columns[0]])
        df.rename(columns={df.columns[0]: 'Despesa'}, inplace=True)
        df = df.dropna(subset=['Despesa'])
        df = df[df['Despesa'].astype(str).str.upper() != 'TOTAL']
        
        meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
        for col in meses:
            if col in df.columns:
                df[col] = df[col].replace('-', 0)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
        # Focando no mês atual de análise
        mes_atual = "JUL"
        
        # Cálculos dos Cards
        total_gasto = df[mes_atual].sum()
        linha_maior = df.loc[df[mes_atual].idxmax()]
        maior_nome = linha_maior['Despesa']
        maior_valor = linha_maior[mes_atual]
        
        dados_resumo = {
            "mes": "Julho",
            "total_gasto": formata_real(total_gasto),
            "maior_despesa": f"{maior_nome} ({formata_real(maior_valor)})"
        }
        
        # Prepara dados para a tabela
        df_tabela = df[['Despesa', mes_atual]].copy()
        df_tabela = df_tabela[df_tabela[mes_atual] > 0]
        df_tabela = df_tabela.sort_values(by=mes_atual, ascending=False)
        
        lista_despesas = []
        for _, row in df_tabela.iterrows():
            lista_despesas.append({
                "nome": row['Despesa'],
                "valor": formata_real(row[mes_atual])
            })
            
        return render_template('index.html', dados=dados_resumo, despesas=lista_despesas)

    except Exception as e:
        # Se houver qualquer outro erro interno, ele nos avisa na tela em vez de quebrar
        return f"<h1>Erro interno do Servidor</h1><p>Falha ao processar os dados: {e}</p>"

# Exporta explicitamente o app para o Vercel
app = app

if __name__ == '__main__':
    app.run(debug=True)