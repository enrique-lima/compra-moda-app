import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.graph_objects as go
import unicodedata, time
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq

# Configurações Visuais e Logo via link web
st.image("https://raw.githubusercontent.com/enrique-lima/compra-moda-app/main/LOGO_TL.png", width=300)

st.markdown(
    """
    <style>
        .main {background-color:#f8f9fa;}
        h1   {color:#004080;}
        .css-1d391kg{font-size:32px;font-weight:bold;color:#004080;}
        .stButton>button{background-color:#004080;color:#fff;font-weight:bold;border-radius:8px;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Previsão de Vendas e Reposição de Estoque")
st.write(
    """
    Este app gera previsão de vendas por **linha OTB** e **cor de produto** para os próximos **6 meses**, 
    ajustada pela tendência de buscas do Google Trends e sazonalidade, além da recomendação de estoque.

    📂 Faça upload de um arquivo Excel com as abas **VENDA** e **ESTOQUE**.
    """
)

# Funções utilitárias
def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        unicodedata.normalize("NFKD", c).encode("ASCII", "ignore").decode("utf-8").strip().lower().replace(" ", "_")
        for c in df.columns
    ]
    return df

def get_trend_uplift(linhas_otb: list[str]) -> dict[str, float]:
    pytrends = TrendReq(hl="pt-BR", tz=360)
    genericos = [
        "acessorios","alpargata","anabela","mocassim","bolsa","bota","cinto","loafer","rasteira",
        "sandalia","sapatilha","scarpin","tenis","meia","meia pata","salto","salto fino",
        "salto normal","sapato tratorado","mule","oxford","papete","peep flat","slide",
    ]
    tendencias: dict[str,float] = {}
    for linha in linhas_otb:
        try:
            termos = [linha.lower()] + genericos
            pytrends.build_payload(termos, timeframe="today 3-m", geo="BR")
            df_trend = pytrends.interest_over_time()
            if not df_trend.empty:
                base = df_trend.get(linha.lower(), pd.Series(index=df_trend.index, data=0)).mean()
                generico = df_trend[genericos].mean(axis=1).mean()
                uplift = ((base + generico) / 2 - 50) / 100
            else:
                uplift = 0
        except Exception:
            uplift = 0
        tendencias[linha] = round(uplift, 3)
        time.sleep(1)
    return tendencias

def forecast_serie(serie: pd.Series, passos: int = 6, sazonalidade: bool = False) -> pd.Series:
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

# Upload do Excel
uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_venda = normalizar_colunas(xls.parse("VENDA"))
    df_estoque = normalizar_colunas(xls.parse("ESTOQUE"))

    required_venda = ["linha_otb", "cor_produto", "qtd_vendida", "ano_venda", "mes_venda", "filial"]
    required_estoque = ["linha", "cor", "filial", "saldo_empresa"]
    if missing := [c for c in required_venda if c not in df_venda.columns]:
        st.error(f"Faltam colunas na aba VENDA: {', '.join(missing)}"); st.stop()
    if missing := [c for c in required_estoque if c not in df_estoque.columns]:
        st.error(f"Faltam colunas na aba ESTOQUE: {', '.join(missing)}"); st.stop()

    meses = {
        "janeiro":1,"fevereiro":2,"marco":3,"abril":4,"maio":5,"junho":6,
        "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
    }
    df_venda["mes_num"] = df_venda["mes_venda"].str.lower().map(meses)
    df_venda = df_venda.dropna(subset=["ano_venda", "mes_num"])
    df_venda["ano_mes"] = pd.to_datetime(df_venda["ano_venda"].astype(int).astype(str)+"-"+df_venda["mes_num"].astype(int).astype(str)+"-01")

    st.sidebar.subheader("⚙️ Ajustes de Forecast")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", 0, 100, 100, 5) / 100
    usar_sazonalidade = st.sidebar.checkbox("Considerar sazonalidade", value=True)

    trend_uplift = get_trend_uplift(df_venda["linha_otb"].dropna().unique().tolist())

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")
    resultado = []

    for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
        serie = grupo.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0)
        prev = forecast_serie(serie, passos=periodos, sazonalidade=usar_sazonalidade)
        ajuste = trend_uplift.get(linha, 0) * peso_google_trends
        prev_adj = (prev * (1 + ajuste)).clip(lower=0)
        estoque_rec = (prev_adj * 2.8).round()
        estoque_atual = df_estoque.loc[
            (df_estoque["linha"] == linha) &
            (df_estoque["cor"] == cor) &
            (df_estoque["filial"] == filial),
            "saldo_empresa"
        ].sum()

        # Cobertura estoque meses
        cobertura = estoque_atual / prev_adj.mean() if prev_adj.mean() > 0 else 0

        # Flag recomendação
        if cobertura > 2.8:
            recomendacao = "Acelerar Venda"
        else:
            recomendacao = "Necessidade Recompra"

        registro = {
            "linha_otb": linha, "cor_produto": cor, "filial": filial, "estoque_atual": estoque_atual,
            "cobertura_estoque_meses": round(cobertura, 2), "recomendacao": recomendacao
        }
        for dt, val, est_rec in zip(prev_adj.index, prev_adj.values, estoque_rec.values):
            registro[f"venda_prevista_{dt.strftime('%Y_%m')}"] = round(val, 0)
            registro[f"estoque_recomendado_{dt.strftime('%Y_%m')}"] = int(est_rec)
        resultado.append(registro)

    df_resultado = pd.DataFrame(resultado)

    # --- FILTRANDO APENAS OS 6 MESES FUTUROS NO OUTPUT FINAL ---
    colunas_prev_estoque = []
    for dt in datas_prev:
        colunas_prev_estoque.append(f"venda_prevista_{dt.strftime('%Y_%m')}")
        colunas_prev_estoque.append(f"estoque_recomendado_{dt.strftime('%Y_%m')}")

    colunas_fixas = ["linha_otb", "cor_produto", "filial", "estoque_atual", "cobertura_estoque_meses", "recomendacao"]

    df_resultado = df_resultado[colunas_fixas + colunas_prev_estoque]

    st.success("Previsão gerada com sucesso!")
    st.dataframe(df_resultado)

    # Gráfico combinado por linha_otb (agregando filiais e cores)
    df_graf = df_resultado.groupby("linha_otb").sum(numeric_only=True).reset_index()
    df_graf["total_venda_prevista"] = df_graf[[c for c in df_graf.columns if c.startswith("venda_prevista_")]].sum(axis=1)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_graf["linha_otb"],
        y=df_graf["total_venda_prevista"],
        name="Venda Prevista 6 meses",
        marker_color="#004080",
    ))
    fig.update_layout(title="Venda Prevista por Linha OTB (6 meses futuros)", xaxis_title="Linha OTB", yaxis_title="Quantidade")
    st.plotly_chart(fig, use_container_width=True)
