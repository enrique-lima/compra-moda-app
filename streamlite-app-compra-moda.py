import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq
import time
import plotly.express as px
from PIL import Image
import os

# --- Logo ---
logo_path = "LOGO_TL.png"
if os.path.exists(logo_path):
    logo_image = Image.open(logo_path)
    st.image(logo_image, width=150)
else:
    st.warning("Logo não encontrado. Por favor, confirme se o arquivo LOGO_TL.png está na pasta do app.")

# --- CSS e título ---
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
Este app permite fazer previsão de vendas por linha OTB e cor de produto, com sugestão de estoque baseada na tendência do Google Trends e histórico de vendas.
✉️ Faça upload de um arquivo Excel com as abas `VENDA` e `ESTOQUE`.
""")

def normalizar_colunas(df):
    df.columns = [
        unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('utf-8')
        .strip().lower().replace(' ', '_')
        for col in df.columns
    ]
    return df

def get_trend_uplift(linhas_otb):
    pytrends = TrendReq(hl='pt-BR', tz=360)
    genericos = [
        'acessorios', 'alpargata', 'anabela', 'mocassim', 'bolsa', 'bota', 'cinto', 'loafer', 'rasteira',
        'sandalia', 'sapatilha', 'scarpin', 'tenis', 'meia', 'meia pata', 'salto', 'salto fino',
        'salto normal', 'sapato tratorado', 'mule', 'oxford', 'papete', 'peep flat', 'slide'
    ]

    tendencias = {}
    for linha in linhas_otb:
        try:
            termos_busca = [linha.lower()] + genericos
            pytrends.build_payload(termos_busca, cat=0, timeframe='today 3-m', geo='BR')
            df_trends = pytrends.interest_over_time()
            if not df_trends.empty:
                media_base = df_trends[linha.lower()] if linha.lower() in df_trends.columns else pd.Series([0])
                media_base = media_base.mean() if not media_base.empty else 0
                media_genericos = df_trends[genericos].mean(axis=1).mean()
                uplift = ((media_base + media_genericos) / 2 - 50) / 100
                tendencias[linha] = round(uplift, 3)
            else:
                pytrends.build_payload(genericos, cat=0, timeframe='today 3-m', geo='BR')
                df_trends = pytrends.interest_over_time()
                media_genericos = df_trends[genericos].mean(axis=1).mean() if not df_trends.empty else 50
                uplift = (media_genericos - 50) / 100
                tendencias[linha] = round(uplift, 3)
            time.sleep(1)
        except Exception:
            tendencias[linha] = 0
    return tendencias

def forecast_serie(serie, passos=3):
    if serie.count() >= 6:
        modelo = ExponentialSmoothing(serie, trend='add', seasonal=None)
        modelo_fit = modelo.fit()
        previsao = modelo_fit.forecast(passos)
    else:
        previsao = pd.Series(
            [serie.mean()] * passos,
            index=pd.date_range(serie.index[-1] + relativedelta(months=1), periods=passos, freq='MS')
        )
    return previsao

uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse('VENDA')
    df_estoque = xls.parse('ESTOQUE')

    df = normalizar_colunas(df)
    df_estoque = normalizar_colunas(df_estoque)

    # Corrigir possíveis erros de digitação
    if 'tamanho_produot' in df.columns:
        df = df.rename(columns={'tamanho_produot': 'tamanho_produto'})

    # Validação básica colunas essenciais
    required_venda_cols = ['linha_otb', 'cor_produto', 'qtd_vendida', 'ano_venda', 'mes_venda']
    missing_cols_venda = [c for c in required_venda_cols if c not in df.columns]
    if missing_cols_venda:
        st.error(f"Faltam colunas no VENDA: {', '.join(missing_cols_venda)}")
        st.stop()

    required_estoque_cols = ['linha', 'cor', 'saldo_empresa']
    missing_cols_estoque = [c for c in required_estoque_cols if c not in df_estoque.columns]
    if missing_cols_estoque:
        st.error(f"Faltam colunas no ESTOQUE: {', '.join(missing_cols_estoque)}")
        st.stop()

    # Preparar datas
    meses_ordem = {
        'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'abril': 4,
        'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
        'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
    }

    df['mes_num'] = df['mes_venda'].str.lower().map(meses_ordem)
    df = df.dropna(subset=['ano_venda', 'mes_num'])
    df['mes_num'] = df['mes_num'].astype(int)
    df['ano_mes'] = pd.to_datetime(df['ano_venda'].astype(int).astype(str) + '-' + df['mes_num'].astype(str) + '-01')

    # Obter linhas OTB únicas e calcular uplift
    linhas_otb_unicas = df['linha_otb'].dropna().unique().tolist()
    with st.spinner('Consultando Google Trends... isso pode demorar um pouco...'):
        trend_uplift = get_trend_uplift(linhas_otb_unicas)

    periodos_forecast = 3
    resultado = []

    # Forecast + ajuste
    for (linha_otb, cor_produto), grupo in df.groupby(['linha_otb', 'cor_produto']):
        serie = grupo.groupby('ano_mes')['qtd_vendida'].sum().sort_index()
        serie = serie.asfreq('MS').fillna(0)
        prev = forecast_serie(serie, passos=periodos_forecast)
        g = trend_uplift.get(linha_otb, 0)
        prev_adj = prev * (1 + g)
        estoque_rec = (prev_adj.mean() * 2.8).round()

        estoque_atual = df_estoque.loc[
            (df_estoque['linha'] == linha_otb) & (df_estoque['cor'] == cor_produto),
            'saldo_empresa'
        ]
        estoque_atual = pd.to_numeric(estoque_atual, errors='coerce').sum()

        registro = {
            'linha_otb': linha_otb,
            'cor_produto': cor_produto,
            'estoque_atual': estoque_atual
        }

        for dt_prev, val in prev_adj.items():
            registro[f'venda_prevista_{dt_prev.strftime("%Y_%m")}'] = round(val, 0)

        registro['estoque_recomendado_total'] = int(estoque_rec)
        resultado.append(registro)

    df_resultado = pd.DataFrame(resultado)

    # Ordenar colunas para exibição
    datas_previstas = pd.date_range(df['ano_mes'].max() + relativedelta(months=1), periods=periodos_forecast, freq='MS')
    meta_cols = ['linha_otb', 'cor_produto', 'estoque_atual'] + [f'venda_prevista_{dt.strftime("%Y_%m")}' for dt in datas_previstas] + ['estoque_recomendado_total']
    df_resultado = df_resultado[meta_cols]

    st.success("Previsão gerada com sucesso!")
    st.dataframe(df_resultado)

    st.subheader("📊 Gráfico de Previsão de Vendas")
    linha_filtrada = st.selectbox("Selecione a linha do produto:", df_resultado['linha_otb'].unique())
    df_grafico = df_resultado[df_resultado['linha_otb'] == linha_filtrada].melt(
        id_vars=['linha_otb', 'cor_produto'],
        value_vars=[f'venda_prevista_{dt.strftime("%Y_%m")}' for dt in datas_previstas],
        var_name='mês',
        value_name='vendas_previstas'
    )
    fig = px.bar(df_grafico, x='mês', y='vendas_previstas', color='cor_produto', barmode='group',
                 labels={'mês': 'Mês', 'vendas_previstas': 'Vendas Previstas'},
                 title=f"Previsão de Vendas para {linha_filtrada}")
    st.plotly_chart(fig)

    # Botão para download do Excel
    output = BytesIO()
    df_resultado.to_excel(output, index=False)
    output.seek(0)
    st.download_button(
        label="📅 Baixar resultados em Excel",
        data=output,
        file_name="forecast_sugestao_compras.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
