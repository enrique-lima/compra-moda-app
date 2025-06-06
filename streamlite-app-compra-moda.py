import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.graph_objects as go
import unicodedata, time
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq
from PIL import Image
import requests
from io import BytesIO as IOBytes

# ────────────────────────────────────────────────────────────────────────────────
# Logo do repositório local (GitHub raw link)
LOGO_URL = "https://raw.githubusercontent.com/enrique-lima/compra-moda-app/9ac980086bec03f84b0546d558f0ef55245193af/LOGO_TL.png"
response = requests.get(LOGO_URL)
img_logo = Image.open(IOBytes(response.content))
st.image(img_logo, width=150)

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

    📄 [Template Excel para upload](https://docs.google.com/spreadsheets/d/1ip_FU9Ah1zjyFhaW6vVUP87I_sosy0r8/edit?usp=drive_link&ouid=105921193969743336299&rtpof=true&sd=true)
    """
)

# ────────────────────────────────────────────────────────────────────────────────
# Funções utilitárias
# ────────────────────────────────────────────────────────────────────────────────

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
    if len(serie.dropna()) < 2 * 12 and sazonalidade:
        # Não há dados suficientes para sazonalidade, ignora e faz forecast simples
        sazonalidade = False

    if sazonalidade:
        modelo = ExponentialSmoothing(serie, trend="add", seasonal="add", seasonal_periods=12)
        prev = modelo.fit().forecast(passos)
    else:
        if serie.count() >= 6:
            modelo = ExponentialSmoothing(serie, trend="add", seasonal=None)
            prev = modelo.fit().forecast(passos)
        else:
            prev = pd.Series(
                [serie.mean()] * passos,
                index=pd.date_range(serie.index[-1] + relativedelta(months=1), periods=passos, freq="MS"),
            )
    return prev

def mes_para_estacao(mes: int) -> str:
    # Mapeamento das estações brasileiras considerando hemisfério sul
    if mes in [12, 1, 2]:
        return "Verão"
    elif mes in [3, 4, 5]:
        return "Outono"
    elif mes in [6, 7, 8]:
        return "Inverno"
    else:
        return "Primavera"

# ────────────────────────────────────────────────────────────────────────────────
# Upload do Excel
# ────────────────────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_venda = normalizar_colunas(xls.parse("VENDA"))
    df_estoque = normalizar_colunas(xls.parse("ESTOQUE"))

    required_venda = ["linha_otb", "cor_produto", "qtd_vendida", "ano_venda", "mes_venda", "filial"]
    required_estoque = ["linha", "cor", "saldo_empresa"]

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
    df_venda["estacao"] = df_venda["mes_num"].apply(mes_para_estacao)

    st.sidebar.subheader("⚙️ Ajustes de Forecast")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", min_value=0, max_value=100, value=100, step=5) / 100
    usar_sazonalidade = st.sidebar.checkbox("Considerar sazonalidade (estações do ano)", value=True)

    trend_uplift = get_trend_uplift(df_venda["linha_otb"].dropna().unique().tolist())

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")

    resultado = []

    # Preparar forecast e estoque por granularidade linha_otb, cor_produto e filial
    for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
        serie = grupo.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0)
        prev = forecast_serie(serie, passos=periodos, sazonalidade=usar_sazonalidade)
        ajuste = trend_uplift.get(linha, 0) * peso_google_trends
        prev_adj = (prev * (1 + ajuste)).clip(lower=0)

        estoque_atual = df_estoque.loc[
            (df_estoque["linha"] == linha) & (df_estoque["cor"] == cor), "saldo_empresa"
        ].sum()

        # Cobertura em meses = estoque_atual dividido pela média mensal prevista
        media_mensal_prev = prev_adj.mean() if prev_adj.mean() > 0 else 1
        cobertura_estoque = estoque_atual / media_mensal_prev

        # Estoque recomendado mensal e total para os próximos 6 meses
        estoque_rec_mensal = prev_adj * 2.8  # cobertura desejada
        estoque_rec_total = estoque_rec_mensal.sum()

        # Flag recomendação
        if cobertura_estoque > 2.8:
            recomendacao = "Acelerar venda"
        else:
            recomendacao = "Necessidade recompra"

        linha_resultado = {
            "linha_otb": linha,
            "cor_produto": cor,
            "filial": filial,
            "estoque_atual": estoque_atual,
            "cobertura_estoque_meses": round(cobertura_estoque, 2),
            "recomendacao_estoque": recomendacao,
        }

        # Adiciona vendas previstas para cada mês
        for i, data in enumerate(datas_prev):
            linha_resultado[f"venda_prevista_{data.strftime('%Y_%m%d')}"] = round(prev_adj[i], 2)
            linha_resultado[f"estoque_recomendado_{data.strftime('%Y_%m%d')}"] = round(estoque_rec_mensal[i], 2)

        # Estoque recomendado total por linha_otb + cor + filial
        linha_resultado["estoque_recomendado_total"] = round(estoque_rec_total, 2)

        resultado.append(linha_resultado)

    df_resultado = pd.DataFrame(resultado)

    st.subheader("📋 Resultado: Forecast e Recomendação de Estoque")
    st.dataframe(df_resultado)

    # ────────── Gráfico Combinado (Linha OTB) ──────────
    st.subheader("📊 Gráfico Combinado: Venda Projetada e Estoque Recomendado (Linha OTB)")

    linhas_sel = st.multiselect(
        "Selecione linhas OTB:", options=df_resultado["linha_otb"].unique(), default=df_resultado["linha_otb"].unique()[:1]
    )

    if linhas_sel:
        df_plot = df_resultado[df_resultado["linha_otb"].isin(linhas_sel)]

        # Melt vendas previstas
        df_vendas = df_plot.melt(
            id_vars=["linha_otb"],
            value_vars=[c for c in df_resultado.columns if c.startswith("venda_prevista_")],
            var_name="mes", value_name="vendas_previstas",
        )
        df_vendas["mes_dt"] = pd.to_datetime(df_vendas["mes"].str.replace("venda_prevista_", "") + "01", format="%Y_%m%d")

        # Agregar estoque recomendado total por linha_otb
        estoque_rec_por_linha = df_plot.groupby("linha_otb")["estoque_recomendado_total"].sum().reset_index()
        estoque_rec_por_linha = estoque_rec_por_linha.loc[estoque_rec_por_linha["linha_otb"].isin(linhas_sel)]

        meses = pd.date_range(df_vendas["mes_dt"].min(), periods=6, freq="MS")
        lista_estoque = []
        for _, row in estoque_rec_por_linha.iterrows():
            for mes in meses:
                lista_estoque.append(
                    {"linha_otb": row["linha_otb"], "mes_dt": mes, "estoque_recomendado": row["estoque_recomendado_total"] / 6}
                )
        df_estoque_plot = pd.DataFrame(lista_estoque)

        df_combined = pd.merge(df_vendas, df_estoque_plot, on=["linha_otb", "mes_dt"], how="left")

        fig = go.Figure()
        for linha in linhas_sel:
            df_linha = df_combined[df_combined["linha_otb"] == linha]

            fig.add_trace(
                go.Bar(
                    x=df_linha["mes_dt"],
                    y=df_linha["estoque_recomendado"],
                    name=f"Estoque Recomendado - {linha}",
                    opacity=0.6,
                )
            )
            fig.add_trace(
                go.Scatter(
                    x=df_linha["mes_dt"],
                    y=df_linha["vendas_previstas"],
                    name=f"Venda Prevista - {linha}",
                    mode="lines+markers",
                    line=dict(width=3),
                )
            )

        fig.update_layout(
            title="Venda Prevista (Linha) e Estoque Recomendado (Colunas) por Linha OTB",
            xaxis_title="Mês",
            yaxis_title="Quantidade",
            barmode="group",
            legend_title="Legenda",
            xaxis=dict(tickformat="%b %Y"),
            height=600,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Selecione ao menos uma linha para visualizar o gráfico.")

else:
    st.info("Aguardando upload do arquivo Excel...")
