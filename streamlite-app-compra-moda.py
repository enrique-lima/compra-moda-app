import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
from PIL import Image
import os

# Tente abrir a imagem do logo de forma segura
logo_path = "LOGO_TL.png"
if os.path.exists(LOGO_TL):
    logo_image = Image.open(LOGO_TL)
    st.image(logo_image, width=150)
else:
    st.warning("Logo não encontrado. Por favor, confirme se o arquivo LOGO_TL.png está na pasta do app.")

# === Estilização via CSS ===
st.markdown("""
    <style>
        .main {
            background-color: #f8f9fa;
        }
        h1 {
            color: #004080;
        }
        .css-1d391kg {
            font-size: 32px;
            font-weight: bold;
            color: #004080;
        }
        .stButton > button {
            background-color: #004080;
            color: white;
            font-weight: bold;
            border-radius: 8px;
        }
    </style>
""", unsafe_allow_html=True)

# === Título e descrição ===
st.title("Previsão de Vendas e Reposição de Estoque")
st.write("""
Este app permite fazer previsão de vendas por linha OTB e cor de produto, com sugestão de estoque com base na tendência do Google Trends e histórico de vendas.

✉️ Basta enviar um arquivo Excel com as abas `VENDA` e `ESTOQUE`.
""")

# === Upload do arquivo ===
uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df = pd.read_excel(xls, 'VENDA')
    df_estoque = pd.read_excel(xls, 'ESTOQUE')

    # Simula forecast simplificado (mock)
    df_resultado = df.groupby(['linha_otb', 'cor_produto'])[['qtd_vendida']].sum().reset_index()
    df_resultado['venda_prevista_mes_1'] = df_resultado['qtd_vendida'] * 1.1
    df_resultado['venda_prevista_mes_2'] = df_resultado['qtd_vendida'] * 1.15
    df_resultado['venda_prevista_mes_3'] = df_resultado['qtd_vendida'] * 1.2
    df_resultado['estoque_recomendado_total'] = (df_resultado[['venda_prevista_mes_1','venda_prevista_mes_2','venda_prevista_mes_3']].mean(axis=1) * 2.8).round()

    st.success("Previsão gerada com sucesso!")
    st.dataframe(df_resultado)

    # Gráfico interativo de forecast
    st.subheader("📊 Gráfico de Previsão de Vendas")
    linha_filtrada = st.selectbox("Selecione a linha do produto:", df_resultado['linha_otb'].unique())
    df_grafico = df_resultado[df_resultado['linha_otb'] == linha_filtrada].melt(
        id_vars=['linha_otb', 'cor_produto'],
        value_vars=['venda_prevista_mes_1', 'venda_prevista_mes_2', 'venda_prevista_mes_3'],
        var_name='mes',
        value_name='vendas_previstas'
    )
    fig = px.bar(df_grafico, x='mes', y='vendas_previstas', color='cor_produto', barmode='group',
                 labels={'mes': 'Mês', 'vendas_previstas': 'Vendas Previstas'},
                 title=f"Previsão de Vendas para {linha_filtrada}")
    st.plotly_chart(fig)

    # Exportar Excel
    output = BytesIO()
    df_resultado.to_excel(output, index=False)
    output.seek(0)
    st.download_button(
        label="📅 Baixar resultados em Excel",
        data=output,
        file_name="forecast_resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
