import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
from PIL import Image
import os

# Logo
logo_path = "LOGO_TL.png"
if os.path.exists(logo_path):
    logo_image = Image.open(logo_path)
    st.image(logo_image, width=150)
else:
    st.warning("Logo não encontrado. Por favor, confirme se o arquivo LOGO_TL.png está na pasta do app.")

# CSS e título
st.markdown("""
    <style>
        .main {background-color: #f8f9fa;}
        h1 {color: #004080;}
        .css-1d391kg {font-size: 32px; font-weight: bold; color: #004080;}
        .stButton > button {
            background-color: #004080; color: white; font-weight: bold; border-radius: 8px;
        }
    </style>
""", unsafe_allow_html=True)

st.title("Previsão de Vendas e Reposição de Estoque")
st.write("""
Este app permite fazer previsão de vendas por linha OTB e cor de produto, com sugestão de estoque com base na tendência do Google Trends e histórico de vendas.
✉️ Basta enviar um arquivo Excel com as abas `VENDA` e `ESTOQUE`.
""")

uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df = pd.read_excel(xls, 'VENDA')
    df_estoque = pd.read_excel(xls, 'ESTOQUE')

    # DEBUG: mostra colunas para ajudar diagnóstico
    st.write("Colunas do DataFrame 'VENDA':", df.columns.tolist())
    st.write(df.head())

    # Define as colunas que você precisa
    col_linha = 'linha_otb'
    col_cor = 'cor'
    col_qtd = 'qtd_vendida'

    # Valida existência das colunas
    faltantes = [col for col in [col_linha, col_cor, col_qtd] if col not in df.columns]
    if faltantes:
        st.error(f"Erro: As seguintes colunas estão faltando no arquivo VENDA: {', '.join(faltantes)}")
    else:
        # Processo normal
        df_resultado = df.groupby([col_linha, col_cor])[[col_qtd]].sum().reset_index()
        df_resultado['venda_prevista_mes_1'] = df_resultado[col_qtd] * 1.1
        df_resultado['venda_prevista_mes_2'] = df_resultado[col_qtd] * 1.15
        df_resultado['venda_prevista_mes_3'] = df_resultado[col_qtd] * 1.2
        df_resultado['estoque_recomendado_total'] = (
            df_resultado[['venda_prevista_mes_1','venda_prevista_mes_2','venda_prevista_mes_3']]
            .mean(axis=1) * 2.8
        ).round()

        st.success("Previsão gerada com sucesso!")
        st.dataframe(df_resultado)

        st.subheader("📊 Gráfico de Previsão de Vendas")
        linha_filtrada = st.selectbox("Selecione a linha do produto:", df_resultado[col_linha].unique())
        df_grafico = df_resultado[df_resultado[col_linha] == linha_filtrada].melt(
            id_vars=[col_linha, col_cor],
            value_vars=['venda_prevista_mes_1', 'venda_prevista_mes_2', 'venda_prevista_mes_3'],
            var_name='mes',
            value_name='vendas_previstas'
        )
        fig = px.bar(df_grafico, x='mes', y='vendas_previstas', color=col_cor, barmode='group',
                    labels={'mes': 'Mês', 'vendas_previstas': 'Vendas Previstas'},
                    title=f"Previsão de Vendas para {linha_filtrada}")
        st.plotly_chart(fig)

        output = BytesIO()
        df_resultado.to_excel(output, index=False)
        output.seek(0)
        st.download_button(
            label="📅 Baixar resultados em Excel",
            data=output,
            file_name="forecast_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
