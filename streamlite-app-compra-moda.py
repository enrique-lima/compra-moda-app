import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.graph_objects as go
import unicodedata
import time
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq
from datetime import datetime

# --- CACHE para upload e parsing do Excel ---
@st.cache_data(show_spinner=False)
def carregar_dados(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    df_venda = normalizar_colunas(xls.parse("VENDA"))
    df_estoque = normalizar_colunas(xls.parse("ESTOQUE"))
    return df_venda, df_estoque

# Função normalizar colunas

def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        unicodedata.normalize("NFKD", c).encode("ASCII", "ignore").decode("utf-8").strip().lower().replace(" ", "_")
        for c in df.columns
    ]
    return df

# --- CACHE para Google Trends (inclui concorrentes) ---
@st.cache_data(show_spinner=False, max_entries=32, ttl=3600*6)
def get_trend_uplift(linhas_otb: tuple[str]) -> tuple[dict[str, float], pd.DataFrame]:
    linhas_otb = list(linhas_otb)
    pytrends = TrendReq(hl="pt-BR", tz=360)

    genericos = [
        "acessorios","alpargata","anabela","mocassim","bolsa","bota","cinto","loafer","rasteira",
        "sandalia","sapatilha","scarpin","tenis","meia","meia pata","salto","salto fino",
        "salto normal","sapato tratorado","mule","oxford","papete","peep flat","slide",
        "sandália spike","salto spike","papete spike"
    ]

    concorrentes = [
        "alexander birman", "schutz", "arezzo", "luiza barcelos", "sidewalk"
    ]

    tendencias: dict[str, float] = {}
    registros = []

    for linha in linhas_otb:
        try:
            termos = [linha.lower()] + genericos + concorrentes
            pytrends.build_payload(termos, timeframe="today 3-m", geo="BR")
            df_trend = pytrends.interest_over_time()
            if not df_trend.empty:
                base_score = df_trend.get(linha.lower(), pd.Series(index=df_trend.index, data=0)).mean()
                generico_score = df_trend[genericos].mean(axis=1).mean()
                concorrente_score = df_trend[concorrentes].mean(axis=1).mean()
                media_geral = (base_score + generico_score + concorrente_score) / 3
                uplift = (media_geral - 50) / 100
            else:
                base_score = generico_score = concorrente_score = uplift = 0
        except Exception:
            base_score = generico_score = concorrente_score = uplift = 0
        tendencias[linha] = round(uplift, 3)
        registros.append({
            "linha_otb": linha,
            "score_linha": round(base_score, 2),
            "score_generico": round(generico_score, 2),
            "score_concorrente": round(concorrente_score, 2),
            "uplift_aplicado": round(uplift, 3)
        })
        time.sleep(1)

    return tendencias, pd.DataFrame(registros)

# --- CACHE para forecast de séries ---
@st.cache_data(show_spinner=False, max_entries=128)
def forecast_serie_cache(serie_values: tuple, serie_index: tuple, passos: int, sazonalidade: bool) -> pd.Series:
    serie = pd.Series(serie_values, index=pd.to_datetime(serie_index))
    if serie.count() >= 24 and sazonalidade:
        modelo = ExponentialSmoothing(serie, trend="add", seasonal="add", seasonal_periods=12)
        prev = modelo.fit().forecast(passos)
    elif serie.count() >= 6:
        modelo = ExponentialSmoothing(serie, trend="add", seasonal=None)
        prev = modelo.fit().forecast(passos)
    else:
        prev = pd.Series(
            [serie.mean()] * passos,
            index=pd.date_range(serie.index[-1] + relativedelta(months=1), periods=passos, freq="MS"),
        )
    return prev.clip(lower=0)

# --- App Interface ---
st.image("https://raw.githubusercontent.com/enrique-lima/compra-moda-app/main/LOGO_TL.png", width=300)
st.title("Previsão de Vendas e Reposição de Estoque")
st.markdown("""
Este app realiza forecast de vendas e recomendações de compra, com análise por Filial, Linha OTB e Cor. Utiliza dados históricos, estoque e tendências de mercado (Google Trends) para prever os próximos 6 meses e apoiar decisões estratégicas.
""")

uploaded_file = st.file_uploader("\U0001F4C2 Faça upload do arquivo Excel", type=["xlsx"])

st.markdown(
    """
    <a href="https://drive.google.com/uc?export=download&id=1ip_FU9Ah1zjyFhaW6vVUP87I_sosy0r8" target="_blank">
        <button style="background-color:#004080;color:white;padding:10px 20px;border:none;border-radius:8px;font-weight:bold;">
            ⬇️ Baixar Template Excel
        </button>
    </a>
    """,
    unsafe_allow_html=True
)

if uploaded_file:
    progresso = st.progress(0)
    status_text = st.empty()

    status_text.text("1/3 - Lendo arquivo Excel...")
    df_venda, df_estoque = carregar_dados(uploaded_file)
    progresso.progress(33)

    status_text.text("2/3 - Normalizando e preparando dados...")
    meses = {"janeiro":1,"fevereiro":2,"marco":3,"abril":4,"maio":5,"junho":6,"julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12}
    df_venda["mes_num"] = df_venda["mes_venda"].str.lower().map(meses)
    df_venda = df_venda.dropna(subset=["ano_venda", "mes_num"])
    df_venda["ano_mes"] = pd.to_datetime(df_venda["ano_venda"].astype(int).astype(str) + "-" + df_venda["mes_num"].astype(int).astype(str) + "-01")
    progresso.progress(66)

    status_text.text("3/3 - Gerando forecast e recomendações...")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", 0, 100, 100, 5) / 100
    usar_sazonalidade = st.sidebar.checkbox("Considerar sazonalidade (verão, inverno, etc)", value=True)

    linhas_otb_unicas = tuple(df_venda["linha_otb"].dropna().unique().tolist())
    trend_uplift, df_trends_raw = get_trend_uplift(linhas_otb_unicas)

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")
    resultado = []

    for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
        grupo_sorted = grupo.sort_values("ano_mes")
        serie_values = tuple(grupo_sorted.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0).values)
        serie_index = tuple(grupo_sorted.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0).index)

        prev = forecast_serie_cache(serie_values, serie_index, passos=periodos, sazonalidade=usar_sazonalidade)
        ajuste = trend_uplift.get(linha, 0) * peso_google_trends
        prev_adj = (prev * (1 + ajuste)).clip(lower=0)
        estoque_rec = (prev_adj * 2.8).round()
        estoque_atual = df_estoque.loc[(df_estoque["linha"] == linha) & (df_estoque["cor"] == cor) & (df_estoque
