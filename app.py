import streamlit as st
import pandas as pd
from datetime import datetime

# --- CONFIGURAÇÕES INICIAIS ---
st.set_page_config(page_title="Dashboard Financeiro 2026", layout="wide")

st.title("💸 Meu Dashboard Financeiro - 2026")

st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117;
        color: #FAFAFA;
    }
    </style>
""", unsafe_allow_html=True)

def formata_real(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

try:
    # --- 1. LENDO O ARQUIVO EXCEL INTEIRO ---
    arquivo_excel = pd.ExcelFile("Controle Financeiro 2026.xlsx")
    
    # --- 2. PROCESSANDO A ABA DE DESPESAS (DADOS) ---
    df = pd.read_excel(arquivo_excel, sheet_name="Dados", skiprows=9)
    
    if df.columns[0].startswith('Unnamed'):
        df = df.drop(columns=[df.columns[0]])
        
    df.rename(columns={df.columns[0]: 'Despesa'}, inplace=True)
    df = df.dropna(subset=['Despesa'])
    df = df[df['Despesa'].astype(str).str.upper() != 'TOTAL']

    meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    for col in meses + ['TOTAL']:
        if col in df.columns:
            df[col] = df[col].replace('-', 0)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # --- FILTRO DE MÊS NO TOPO DA TELA ---
    st.write("---")
    mes_atual_idx = datetime.now().month - 1 
    mes_selecionado = st.selectbox("📅 Selecione o mês para análise:", meses, index=mes_atual_idx)
    idx_selecionado = meses.index(mes_selecionado)

    # --- CÁLCULOS DOS CARDS DE DESPESA ---
    gasto_mes_atual = df[mes_selecionado].sum()
    linha_maior_despesa = df.loc[df[mes_selecionado].idxmax()]
    nome_maior_despesa = linha_maior_despesa['Despesa']
    valor_maior_despesa = linha_maior_despesa[mes_selecionado]

    if idx_selecionado > 0:
        mes_anterior = meses[idx_selecionado - 1]
        gasto_mes_anterior = df[mes_anterior].sum()
        diferenca_reais = gasto_mes_atual - gasto_mes_anterior
        
        if gasto_mes_anterior > 0:
            diferenca_pct = (diferenca_reais / gasto_mes_anterior) * 100
            texto_delta = f"{diferenca_pct:.1f}% em relação a {mes_anterior}"
        else:
            texto_delta = f"Aumento em relação a {mes_anterior}"
    else:
        gasto_mes_anterior = 0
        texto_delta = "Sem mês anterior (JAN)"

    # --- EXIBINDO OS CARDS DE DESPESA ---
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Gasto no Mês", formata_real(gasto_mes_atual), delta=texto_delta, delta_color="inverse")
    with col2:
        st.metric("Despesa Mais Cara", nome_maior_despesa, delta=formata_real(valor_maior_despesa), delta_color="off")
    with col3:
        st.metric("Média Mensal do Ano", formata_real(df[meses].sum().mean()))

    st.write("---")

    # --- TABELA E GRÁFICOS DE DESPESAS ---
    st.subheader(f"🗂️ Detalhamento de {mes_selecionado}")
    
    df_exibicao = df[['Despesa', mes_selecionado]].copy()
    df_exibicao = df_exibicao.sort_values(by=mes_selecionado, ascending=False)
    
    st.dataframe(
        df_exibicao, 
        column_config={mes_selecionado: st.column_config.NumberColumn("Valor", format="R$ %.2f")}, 
        use_container_width=True, 
        hide_index=True
    )
    
    col_esquerda, col_direita = st.columns(2)
    with col_esquerda:
        st.subheader("💰 Top Despesas do Ano (Total)")
        df_totais = df[['Despesa', 'TOTAL']].sort_values(by='TOTAL', ascending=False).head(10)
        st.bar_chart(df_totais.set_index('Despesa'), color="#FF4B4B")
            
    with col_direita:
        st.subheader("📈 Evolução Anual da Despesa")
        despesa_escolhida = st.selectbox("Escolha a Despesa para detalhar:", df['Despesa'].unique())
        dados_despesa = df[df['Despesa'] == despesa_escolhida][meses].T
        st.line_chart(dados_despesa, color="#FF9800") # Cor laranja para diferenciar dos investimentos

    # ==========================================
    # --- 3. SEÇÃO DE INVESTIMENTOS (NOVIDADE) ---
    # ==========================================
    st.write("---")
    st.header("🌱 Seu Patrimônio e Investimentos")
    
    # Procura pela aba que tenha "Invest" no nome para evitar erros de digitação/espaços
    nome_aba_invest = [aba for aba in arquivo_excel.sheet_names if "INVEST" in aba.upper()]
    
    if nome_aba_invest:
        df_invest = pd.read_excel(arquivo_excel, sheet_name=nome_aba_invest[0])
        
        # Procura a linha onde a primeira coluna diz "Investimento"
        linha_invest = df_invest[df_invest.iloc[:, 0].astype(str).str.contains('Invest', case=False, na=False)]
        
        if not linha_invest.empty:
            # Pega apenas os valores dos meses e converte para número
            valores_mes = linha_invest[meses].iloc[0].replace('-', 0).fillna(0).astype(float)
            
            # Calcula o acumulado (Soma progressiva: Jan, Jan+Fev, Jan+Fev+Mar...)
            acumulado = valores_mes.cumsum()
            total_investido = acumulado['DEZ'] # O total no fim do ano
            
            col_inv1, col_inv2 = st.columns([1, 2])
            
            with col_inv1:
                st.metric("Total Projetado no Ano", formata_real(total_investido))
                st.info("💡 Continue assim! O gráfico ao lado mostra o seu dinheiro se acumulando como uma bola de neve ao longo dos meses.")
                
            with col_inv2:
                st.area_chart(acumulado, color="#00C853") # Verde para crescimento
        else:
            st.warning("Não encontrei a linha 'Investimento' na sua aba de investimentos.")
    else:
        st.warning("Não encontrei uma aba de Investimentos na sua planilha.")

except FileNotFoundError:
    st.error("Arquivo 'Controle Financeiro 2026.xlsx' não encontrado.")
except Exception as e:
    st.error(f"Erro inesperado: {e}")