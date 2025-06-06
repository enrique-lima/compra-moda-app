import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
import unicodedata, time
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq

# ────────────────────────────────────────────────────────────────────────────────
# Logo local (supondo que está na pasta do app)
from PIL import Image
logo = Image.open("LOGO_TL.png")
st.image(logo, width=150)

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
    ajustada pela tendência de buscas do Google Trends, além da recomendação de estoque.

    📂 Faça upload de um arquivo Excel com as abas **VENDA** e **ESTOQUE**.

    📄 [Template Excel para Upload](https://docs.google.com/spreadsheets/d/1ip_FU9Ah1zjyFhaW6vVUP87I_sosy0r8/edit?usp=drive_link&ouid=105921193969743336299&rtpof=true&sd=true)
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
    return prev.clip(lower=0)

uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_venda = normalizar_colunas(xls.parse("VENDA"))
    df_estoque = normalizar_colunas(xls.parse("ESTOQUE"))

    required_venda = ["linha_otb", "cor_produto", "qtd_vendida", "ano_venda", "mes_venda", "filial"]
    required_estoque = ["linha", "cor", "saldo_empresa", "filial"]
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
    df_venda["ano_mes"] = pd.to_datetime(df_venda["ano_venda"].astype(int).astype(str) + "-" + df_venda["mes_num"].astype(int).astype(str) + "-01")

    st.sidebar.subheader("⚙️ Ajustes de Forecast")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", min_value=0, max_value=100, value=100, step=5) / 100
    usar_sazonalidade = st.sidebar.checkbox("Incluir sazonalidade (12 meses)", value=True)

    trend_uplift = get_trend_uplift(df_venda["linha_otb"].dropna().unique().tolist())

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")
    resultado = []

    for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
        serie = grupo.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0)
        prev = forecast_serie(serie, passos=periodos, sazonalidade=usar_sazonalidade)
        ajuste = trend_uplift.get(linha, 0) * peso_google_trends
        prev_adj = (prev * (1 + ajuste)).clip(lower=0)

        estoque_rec_mes = prev_adj * 2.8

        estoque_atual = df_estoque.loc[
            (df_estoque["linha"] == linha) &
            (df_estoque["cor"] == cor) &
            (df_estoque["filial"] == filial),
            "saldo_empresa"
        ].sum()

        media_venda = prev_adj.mean() if prev_adj.mean() > 0 else 1
        cobertura = estoque_atual / media_venda

        if cobertura > 2.8:
            recomendacao = "Acelerar venda"
        else:
            recomendacao = "Necessidade de recompra"

        registro = {
            "linha_otb": linha,
            "cor_produto": cor,
            "filial": filial,
            "estoque_atual": estoque_atual,
            "cobertura_estoque_meses": round(cobertura, 2),
            "recomendacao_estoque": recomendacao,
        }

        for dt, val in prev_adj.items():
            registro[f"venda_prevista_{dt.strftime('%Y_%m')}"] = round(val, 0)
            registro[f"estoque_recomendado_{dt.strftime('%Y_%m')}"] = round(val * 2.8, 0)

        registro["estoque_recomendado_total"] = int(estoque_rec_mes.mean())

        resultado.append(registro)

    df_resultado = pd.DataFrame(resultado)

    colunas = [
        "linha_otb",
        "cor_produto",
        "filial",
        "estoque_atual",
        "cobertura_estoque_meses",
        "recomendacao_estoque",
    ] + [f"venda_prevista_{d.strftime('%Y_%m')}" for d in datas_prev] + [f"estoque_recomendado_{d.strftime('%Y_%m')}" for d in datas_prev] + ["estoque_recomendado_total"]

    df_resultado = df_resultado[colunas]

    st.success("Previsão gerada com sucesso!")
    st.dataframe(df_resultado)

    st.subheader("📊 Gráfico de Previsão de Vendas (Coluna Empilhada)")
    linhas_sel = st.multiselect(
        "Selecione linhas OTB:", options=df_resultado["linha_otb"].unique(), default=df_resultado["linha_otb"].unique()[:1]
    )
    if linhas_sel:
        df_plot = df_resultado[df_resultado["linha_otb"].isin(linhas_sel)].melt(
            id_vars=["linha_otb", "cor_produto", "filial"],
            value_vars=[c for c in df_resultado.columns if c.startswith("venda_prevista_")],
            var_name="mês", value_name="vendas_previstas"
        )
        df_plot["mês_ord"] = pd.to_datetime(df_plot["mês"].str.replace("venda_prevista_", "") + "-01")
        fig = px.bar(df_plot, x="mês_ord", y="vendas_previstas", color="linha_otb", barmode="stack", title="Forecast Vendas")
        st.plotly_chart(fig)

    def to_excel(df: pd.DataFrame) -> bytes:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Previsao")
            writer.save()
        return output.getvalue()

    st.download_button(
        label="📥 Baixar previsão em Excel",
        data=to_excel(df_resultado),
        file_name="previsao_estoque_venda.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
