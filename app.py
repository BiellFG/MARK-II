from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

def formata_real(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

@app.route('/')
def home():
    try:
        # 1. Lê a planilha real
        df = pd.read_excel("Controle Financeiro 2026.xlsx", sheet_name="Dados", skiprows=9)
        
        # 2. Limpa os dados
        if df.columns[0].startswith('Unnamed'):
            df = df.drop(columns=[df.columns[0]])
        df.rename(columns={df.columns[0]: 'Despesa'}, inplace=True)
        df = df.dropna(subset=['Despesa'])
        df = df[df['Despesa'].astype(str).str.upper() != 'TOTAL']
        
        # 3. Trata os números
        meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
        for col in meses:
            if col in df.columns:
                df[col] = df[col].replace('-', 0)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
        # 4. Focando no mês atual (Exemplo: Julho / JUL)
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
        
        # 5. Prepara os dados para a Tabela HTML
        df_tabela = df[['Despesa', mes_atual]].copy()
        df_tabela = df_tabela[df_tabela[mes_atual] > 0] # Remove o que for zero
        df_tabela = df_tabela.sort_values(by=mes_atual, ascending=False)
        
        # Converte para um formato que o HTML entenda
        lista_despesas = []
        for _, row in df_tabela.iterrows():
            lista_despesas.append({
                "nome": row['Despesa'],
                "valor": formata_real(row[mes_atual])
            })
            
        return render_template('index.html', dados=dados_resumo, despesas=lista_despesas)

    except Exception as e:
        return f"<h1>Erro ao ler a planilha:</h1><p>{e}</p>"

app = app

if __name__ == '__main__':
    app.run(debug=True)