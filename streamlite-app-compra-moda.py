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

# Função normalizar colunas (sem cache pois é leve)
def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        unicodedata.normalize("NFKD", c).encode("ASCII", "ignore").decode("utf-8").strip().lower().replace(" ", "_")
        for c in df.columns
    ]
    return df

# --- CACHE para Google Trends ---
@st.cache_data(show_spinner=False, max_entries=32, ttl=3600*6)  # cache por 6 horas
def get_trend_uplift(linhas_otb: tuple[str]) -> dict[str, float]:
    linhas_otb = list(linhas_otb)
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

# --- App ---
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

uploaded_file = st.file_uploader("📂 Faça upload do arquivo Excel", type=["xlsx"])

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
    with st.spinner("Carregando dados e normalizando..."):
        df_venda, df_estoque = carregar_dados(uploaded_file)

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

    st.sidebar.subheader("⚙️ Ajustes de Forecast")
    peso_google_trends = st.sidebar.slider("Peso do ajuste Google Trends (%)", 0, 100, 100, 5) / 100
    usar_sazonalidade = st.sidebar.checkbox("Considerar sazonalidade", value=True)

    linhas_otb_unicas = tuple(df_venda["linha_otb"].dropna().unique().tolist())
    with st.spinner("Consultando Google Trends..."):
        trend_uplift = get_trend_uplift(linhas_otb_unicas)

    periodos = 6
    datas_prev = pd.date_range(df_venda["ano_mes"].max() + relativedelta(months=1), periods=periodos, freq="MS")
    resultado = []

    with st.spinner("Gerando previsões..."):
        for (linha, cor, filial), grupo in df_venda.groupby(["linha_otb", "cor_produto", "filial"]):
            grupo_sorted = grupo.sort_values("ano_mes")
            serie_values = tuple(grupo_sorted.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0).values)
            serie_index = tuple(grupo_sorted.groupby("ano_mes")["qtd_vendida"].sum().sort_index().asfreq("MS", fill_value=0).index)
            
            prev = forecast_serie_cache(serie_values, serie_index, passos=periodos, sazonalidade=usar_sazonalidade)
            ajuste = trend_uplift.get(linha, 0) * peso_google_trends
            prev_adj = (prev * (1 + ajuste)).clip(lower=0)
            estoque_rec = (prev_adj * 2.8).round()
            estoque_atual = df_estoque.loc[
                (df_estoque["linha"] == linha) &
                (df_estoque["cor"] == cor) &
                (df_estoque["filial"] == filial),
                "saldo_empresa"
            ].sum()

            cobertura = estoque_atual / prev_adj.mean() if prev_adj.mean() > 0 else 0

            recomendacao = "Acelerar Venda" if cobertura > 2.8 else "Necessidade Recompra"

            registro = {
                "linha_otb": linha, "cor_produto": cor, "filial": filial, "estoque_atual": estoque_atual,
                "cobertura_estoque_meses": round(cobertura, 2), "recomendacao": recomendacao
            }
            for dt, val, est_rec in zip(prev_adj.index, prev_adj.values, estoque_rec.values):
                registro[f"venda_prevista_{dt.strftime('%Y_%m')}"] = round(val, 0)
                registro[f"estoque_recomendado_{dt.strftime('%Y_%m')}"] = int(est_rec)
            resultado.append(registro)

    df_resultado = pd.DataFrame(resultado)

    colunas_prev_estoque = []
    for dt in datas_prev:
        colunas_prev_estoque.append(f"venda_prevista_{dt.strftime('%Y_%m')}")
        colunas_prev_estoque.append(f"estoque_recomendado_{dt.strftime('%Y_%m')}")

    colunas_fixas = ["linha_otb", "cor_produto", "filial", "estoque_atual", "cobertura_estoque_meses", "recomendacao"]

    df_resultado = df_resultado[colunas_fixas + colunas_prev_estoque]

    st.success("Previsão gerada com sucesso!")
    # Pré-visualização do Excel gerado
    st.subheader("📋 Pré-visualização do Resultado")
    st.dataframe(df_resultado.head(50), use_container_width=True)

    # Exportar Excel para download
     output = BytesIO()
     with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df_resultado.to_excel(writer, index=False, sheet_name='Resumo_Previsao')
    output.seek(0)

    st.download_button(
    label="📥 Baixar Excel com Previsões",
    data=output,
    file_name="previsao_estoque.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


    # Filtro multiselect acima do gráfico com limite opcional (exemplo max 10)
    linhas_otb_disponiveis = df_resultado["linha_otb"].unique()
    linhas_otb_selecionadas = st.multiselect(
        "Selecione as linhas OTB para visualizar no gráfico (máx 10)",
        options=df_resultado["linha_otb"].unique(),
        default=df_resultado["linha_otb"].unique()[:10],  # garante até 10 selecionados por padrão
        max_selections=10,
    )

    df_filtrado = df_resultado[df_resultado["linha_otb"].isin(linhas_otb_selecionadas)]

    venda_cols = [c for c in df_filtrado.columns if c.startswith("venda_prevista_")]
    estoque_cols = [c for c in df_filtrado.columns if c.startswith("estoque_recomendado_")]

    df_venda_long = (
        df_filtrado[["linha_otb"] + venda_cols]
        .melt(id_vars=["linha_otb"], value_vars=venda_cols, var_name="mes", value_name="venda_prevista")
    )
    df_venda_long["mes"] = pd.to_datetime(df_venda_long["mes"].str.replace("venda_prevista_", "") + "_01", format="%Y_%m_%d")

    df_estoque_long = (
        df_filtrado[["linha_otb"] + estoque_cols]
        .melt(id_vars=["linha_otb"], value_vars=estoque_cols, var_name="mes", value_name="estoque_recomendado")
    )
    df_estoque_long["mes"] = pd.to_datetime(df_estoque_long["mes"].str.replace("estoque_recomendado_", "") + "_01", format="%Y_%m_%d")

    df_plot = pd.merge(df_venda_long, df_estoque_long, on=["linha_otb", "mes"])

    hoje = datetime.today().replace(day=1)
    fim_periodo = hoje + relativedelta(months=6)
    df_plot = df_plot[(df_plot["mes"] >= hoje) & (df_plot["mes"] < fim_periodo)]

    fig = go.Figure()

    linhas = df_plot["linha_otb"].unique()

    for i, linha in enumerate(linhas):
        df_linha = df_plot[df_plot["linha_otb"] == linha].sort_values("mes")
        fig.add_trace(go.Scatter(
            x=df_linha["mes"],
            y=df_linha["venda_prevista"],
            name=f"Venda Prevista - {linha}",
            mode="lines+markers",
            line=dict(color=f"rgba({(i*30)%255},{(i*70)%255},{(i*110)%255},1)"),
            yaxis="y1"
        ))
        fig.add_trace(go.Bar(
            x=df_linha["mes"],
            y=df_linha["estoque_recomendado"],
            name=f"Estoque Recomendado - {linha}",
            opacity=0.6,
            offsetgroup=i,
            yaxis="y2"
        ))

    fig.update_layout(
        title="Previsão de Vendas e Estoque Recomendado por Linha OTB",
        xaxis_title="Mês",
        yaxis=dict(title="Venda Prevista", side="left"),
        yaxis2=dict(title="Estoque Recomendado", overlaying="y", side="right"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        barmode="group",
        template="plotly_white",
        height=600,
    )

    st.plotly_chart(fig, use_container_width=True)
