import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
import unicodedata, time
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq

from PIL import Image

# Configurações visuais e logo
logo_path = "LOGO_TL.png"
st.image(logo_path, width=150)

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

st.title("Previsão de Vendas e Estoque com Sazonalidade por Estações")
st.write(
    """
    Forecast considerando as estações do ano (verão, outono, inverno, primavera), 
    linha OTB, cor de produto e filial para os próximos 6 meses, ajustado pela tendência do Google Trends.
    """
)

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

def mes_para_estacao(mes: int) -> str:
    if mes in [12, 1, 2]:
        return "verao"
    elif mes in [3, 4, 5]:
        return "outono"
    elif mes in [6, 7, 8]:
        return "inverno"
    elif mes in [9, 10, 11]:
        return "primavera"
    else:
        return "desconhecida"

def forecast_com_estacao(serie: pd.Series, passos: int = 6) -> pd.Series:
    """
    Forecast com Exponential Smoothing considerando sazonalidade trimestral (4 estações).
    Se houver dados insuficientes, fallback para modelo sem sazonalidade.
    """
    if serie.count() >= 12:
        modelo = ExponentialSmoothing(serie, trend="add", seasonal="add", seasonal_periods=4)
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

uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_venda = normalizar_colunas(xls.parse("VENDA"))
    df_estoque = normalizar_colunas(xls.parse("ESTOQUE"))

    required_venda = ["linha_otb", "cor_produto", "qtd_vendida", "ano_venda", "mes_venda", "filial"]
    required_estoque = ["linha", "cor", "filial", "saldo_empresa"]
    if missing := [c for c in required_venda if c not in df_venda.columns]:
        st.error(f"Faltam colunas na aba VENDA: {', '.join(missing)}")
        st.stop()
    if missing := [c for c in required_estoque if c not in df_estoque.columns]:
        st.error(f"Faltam colunas na aba ESTOQUE: {', '.join(missing)}")
        st.stop()

    meses = {
        "janeiro":1,"fevereiro":2,"marco":3,"abril":4,"maio":5,"junho":6,
        "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
    }
    df_venda["mes_num"] = df_venda["mes_venda"].str.lower().map(meses)
    df_venda = df_venda.dropna(subset=["ano_venda", "mes_num"])
    df_venda["ano_mes"] = pd.to_datetime(df_venda["ano_venda"].astype(int).astype(str) + "-" + df_venda["mes_num"].astype(int).astype(str) + "-01")

    # Coluna estação
    df_venda["estacao"] = df_venda["mes_num"].apply(mes_para_estacao)

    st.sidebar.subheader("⚙️ Ajustes de Forecast")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", 0, 100, 100, 5) / 100

    st.info("Consultando Google Trends… aguarde ~1 minuto se houver muitas linhas.")
    trend_uplift = get_trend_uplift(df_venda["linha_otb"].dropna().unique().tolist())

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")
    resultado = []

    for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
        # Sumarizar venda mensal
        serie = grupo.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0)

        prev = forecast_com_estacao(serie, passos=periodos)
        ajuste = trend_uplift.get(linha, 0) * peso_google_trends
        prev_adj = (prev * (1 + ajuste)).clip(lower=0)

        estoque_atual = df_estoque.loc[
            (df_estoque["linha"] == linha) &
            (df_estoque["cor"] == cor) &
            (df_estoque["filial"] == filial),
            "saldo_empresa"
        ].sum()

        estoque_mensal = []
        cobertura_meses = 2.8
        for i in range(periodos):
            estoque_esperado = estoque_atual + estoque_mensal[-1] if i > 0 else estoque_atual
            estoque_esperado = estoque_esperado + prev_adj.iloc[i] * cobertura_meses
            estoque_mensal.append(round(estoque_esperado, 0))

        registro = {
            "linha_otb": linha,
            "cor_produto": cor,
            "filial": filial,
            "estoque_atual": estoque_atual,
        }
        for dt, val in prev_adj.items():
            registro[f"venda_prevista_{dt.strftime('%Y_%m')}"] = round(val, 0)
        for i, dt in enumerate(datas_prev):
            registro[f"estoque_esperado_{dt.strftime('%Y_%m')}"] = estoque_mensal[i]

        registro["estoque_recomendado_total"] = int((prev_adj.mean() * cobertura_meses).round())
        resultado.append(registro)

    df_resultado = pd.DataFrame(resultado)

    colunas = (
        ["linha_otb", "cor_produto", "filial", "estoque_atual"] +
        [f"venda_prevista_{d.strftime('%Y_%m')}" for d in datas_prev] +
        [f"estoque_esperado_{d.strftime('%Y_%m')}" for d in datas_prev] +
        ["estoque_recomendado_total"]
    )
    df_resultado = df_resultado[colunas]

    st.success("Previsão gerada com sazonalidade por estação!")
    st.dataframe(df_resultado)

    # Gráfico
    st.subheader("📊 Gráfico de Previsão de Vendas")
    linhas_sel = st.multiselect(
        "Selecione linhas OTB:", options=df_resultado["linha_otb"].unique(), default=df_resultado["linha_otb"].unique()[:1]
    )
    if linhas_sel:
        df_plot = df_resultado[df_resultado["linha_otb"].isin(linhas_sel)].melt(
            id_vars=["linha_otb", "cor_produto", "filial"],
            value_vars=[c for c in df_resultado.columns if c.startswith("venda_prevista_")],
            var_name="mês", value_name="vendas_previstas"
        )
        df_plot["mês_ord"] = pd.to_datetime(df_plot["mês"].str.replace("venda_prevista_", "") + "01", format="%Y_%m%d")
        df_plot = df_plot.sort_values("mês_ord")

        fig = px.bar(
            df_plot,
            x="mês",
            y="vendas_previstas",
            color="cor_produto",
            barmode="stack",
            facet_col="linha_otb",
            category_orders={"mês": df_plot["mês"].unique()},
            labels={"mês": "Mês", "vendas_previstas": "Vendas Previstas", "cor_produto": "Cor"},
            title="Previsão de Vendas • Barras Empilhadas por Cor",
            height=600
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
        file_name="forecast_sazonal_estacoes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
