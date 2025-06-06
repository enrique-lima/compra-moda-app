import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
import unicodedata, time
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq

# ──────────────────────────────
# Visual & Logo (local)
# ──────────────────────────────
from PIL import Image
logo = Image.open("LOGO_TL.png")
st.image(logo, width=300)

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
    Este app gera previsão de vendas por **linha OTB**, **cor de produto** e **filial** para os próximos **6 meses**, 
    ajustada pela tendência do Google Trends e sazonalidade (estações do ano), além da recomendação mensal de estoque.

    📂 Faça upload de um arquivo Excel com as abas **VENDA** e **ESTOQUE**.
    """
)

# ──────────────────────────────
# Funções utilitárias
# ──────────────────────────────
def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        unicodedata.normalize("NFKD", c).encode("ASCII", "ignore").decode("utf-8").strip().lower().replace(" ", "_")
        for c in df.columns
    ]
    return df

@st.cache_data(show_spinner=False)
def carregar_dados(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    df_venda = normalizar_colunas(xls.parse("VENDA"))
    df_estoque = normalizar_colunas(xls.parse("ESTOQUE"))
    meses = {
        "janeiro":1,"fevereiro":2,"marco":3,"abril":4,"maio":5,"junho":6,
        "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
    }
    df_venda["mes_num"] = df_venda["mes_venda"].str.lower().map(meses)
    df_venda = df_venda.dropna(subset=["ano_venda", "mes_num"])
    df_venda["ano_mes"] = pd.to_datetime(df_venda["ano_venda"].astype(int).astype(str)+"-"+df_venda["mes_num"].astype(int).astype(str)+"-01")
    return df_venda, df_estoque

@st.cache_data(show_spinner=False)
def calcular_trends(linhas_otb: list[str]) -> dict[str, float]:
    pytrends = TrendReq(hl="pt-BR", tz=360)
    genericos = [
        "acessorios","alpargata","anabela","mocassim","bolsa","bota","cinto","loafer","rasteira",
        "sandalia","sapatilha","scarpin","tenis","meia","meia pata","salto","salto fino",
        "salto normal","sapato tratorado","mule","oxford","papete","peep flat","slide",
    ]
    tendencias = {}
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
        time.sleep(1)  # evita bloqueio
    return tendencias

def forecast_serie(serie: pd.Series, passos: int = 6, sazonalidade: bool = False) -> pd.Series:
    if serie.count() >= 12 and sazonalidade:
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
    return prev

def obter_estacao(dt):
    mes = dt.month
    if mes in [12, 1, 2]:
        return "verao"
    elif mes in [3, 4, 5]:
        return "outono"
    elif mes in [6, 7, 8]:
        return "inverno"
    else:
        return "primavera"

# ──────────────────────────────
# Upload e processamento
# ──────────────────────────────
uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    df_venda, df_estoque = carregar_dados(uploaded_file)

    required_venda = ["linha_otb", "cor_produto", "qtd_vendida", "ano_venda", "mes_venda", "filial"]
    required_estoque = ["linha", "cor", "saldo_empresa"]
    if missing := [c for c in required_venda if c not in df_venda.columns]:
        st.error(f"Faltam colunas na aba VENDA: {', '.join(missing)}"); st.stop()
    if missing := [c for c in required_estoque if c not in df_estoque.columns]:
        st.error(f"Faltam colunas na aba ESTOQUE: {', '.join(missing)}"); st.stop()

    st.sidebar.subheader("⚙️ Ajustes de Forecast")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", min_value=0, max_value=100, value=100, step=5) / 100
    usar_sazonalidade = st.sidebar.checkbox("Incluir sazonalidade (estações do ano)", value=True)

    linhas_otb = df_venda["linha_otb"].dropna().unique().tolist()
    trend_uplift = calcular_trends(linhas_otb)

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")
    resultado = []

    for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
        serie = grupo.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0)

        # Ajustar série para refletir a estação do ano (pode ser feito como variável externa em modelo mais avançado)
        # Aqui usamos sazonalidade no modelo se checkbox ativo e dados suficientes
        prev = forecast_serie(serie, passos=periodos, sazonalidade=usar_sazonalidade)
        
        ajuste = trend_uplift.get(linha, 0) * peso_google_trends
        prev_adj = (prev * (1 + ajuste)).clip(lower=0)

        estoque_atual = df_estoque.loc[
            (df_estoque["linha"] == linha) & (df_estoque["cor"] == cor), "saldo_empresa"
        ].sum()

        # Estoque recomendado mensal com cobertura ideal 2.8 meses, calculado mês a mês
        estoque_recomendado = prev_adj * 2.8

        registro = {
            "linha_otb": linha,
            "cor_produto": cor,
            "filial": filial,
            "estoque_atual": estoque_atual,
        }
        for dt, val_prev, val_estoque in zip(datas_prev, prev_adj, estoque_recomendado):
            registro[f"venda_prevista_{dt.strftime('%Y_%m')}"] = round(val_prev, 0)
            registro[f"estoque_recomendado_{dt.strftime('%Y_%m')}"] = round(val_estoque, 0)
        
        resultado.append(registro)

    df_resultado = pd.DataFrame(resultado)

    colunas = ["linha_otb","cor_produto","filial","estoque_atual"] + \
        [f"venda_prevista_{d.strftime('%Y_%m')}" for d in datas_prev] + \
        [f"estoque_recomendado_{d.strftime('%Y_%m')}" for d in datas_prev]

    df_resultado = df_resultado[colunas]

    st.success("Previsão gerada com sucesso!")
    st.dataframe(df_resultado)

    st.subheader("📊 Gráfico de Previsão de Vendas (Coluna Empilhada)")
    linhas_sel = st.multiselect(
        "Selecione linhas OTB:", options=df_resultado["linha_otb"].unique(), default=df_resultado["linha_otb"].unique()[:1]
    )
    if linhas_sel:
        df_plot = df_resultado[df_resultado["linha_otb"].isin(linhas_sel)].melt(
            id_vars=["linha_otb","cor_produto","filial"],
            value_vars=[c for c in df_resultado.columns if c.startswith("venda_prevista_")],
            var_name="mês", value_name="vendas_previstas"
        )
        df_plot["mês_ord"] = pd.to_datetime(df_plot["mês"].str.replace("venda_prevista_","")+"01", format="%Y_%m%d")
        df_plot = df_plot.sort_values("mês_ord")

        fig = px.bar(
            df_plot, x="mês", y="vendas_previstas", color="cor_produto", barmode="stack",
            facet_col="linha_otb", category_orders={"mês": df_plot["mês"].unique()},
            labels={"mês":"Mês","vendas_previstas":"Vendas Previstas","cor_produto":"Cor"},
            title="Previsão de Vendas • Barras Empilhadas por Cor", height=600
        )
        fig.update_xaxes(tickangle=45)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Selecione ao menos uma linha para visualizar o gráfico.")

    # Download Excel
    buffer = BytesIO()
    df_resultado.to_excel(buffer, index=False)
    buffer.seek(0)
    st.download_button(
        "📅 Baixar Excel de Resultados",
        data=buffer,
        file_name="forecast_sugestao_compras.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
