if uploaded_file:
    # ... seu código até criação do df_resultado ...

    st.success("Previsão gerada com sucesso!")
    st.dataframe(df_resultado)

    # --- Filtro multiseletor para Linha OTB acima do gráfico ---
    linhas_otb_disponiveis = df_resultado["linha_otb"].unique()
    linhas_otb_selecionadas = st.multiselect(
        "Selecione as linhas OTB para visualizar no gráfico",
        options=linhas_otb_disponiveis,
        default=linhas_otb_disponiveis.tolist()
    )

    # Filtra o dataframe conforme seleção
    df_filtrado = df_resultado[df_resultado["linha_otb"].isin(linhas_otb_selecionadas)]

    # Extrai colunas de venda prevista e estoque recomendado
    venda_cols = [c for c in df_filtrado.columns if c.startswith("venda_prevista_")]
    estoque_cols = [c for c in df_filtrado.columns if c.startswith("estoque_recomendado_")]

    # Transforma os dados para formato long para plotagem
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

    # Merge para juntar venda e estoque por linha_otb e mês
    df_plot = pd.merge(df_venda_long, df_estoque_long, on=["linha_otb", "mes"])

    # Filtra para os próximos 6 meses a partir do mês atual
    from datetime import datetime
    hoje = datetime.today().replace(day=1)
    fim_periodo = hoje + relativedelta(months=6)
    df_plot = df_plot[(df_plot["mes"] >= hoje) & (df_plot["mes"] < fim_periodo)]

    # Plot com Plotly - gráfico combinado linha + barra
    import plotly.graph_objects as go

    fig = go.Figure()

    for linha in linhas_otb_selecionadas:
        df_linha = df_plot[df_plot["linha_otb"] == linha]

        fig.add_trace(go.Bar(
            x=df_linha["mes"],
            y=df_linha["estoque_recomendado"],
            name=f"Estoque Recomendado - {linha}",
            opacity=0.6,
            offsetgroup=linha,
        ))

        fig.add_trace(go.Scatter(
            x=df_linha["mes"],
            y=df_linha["venda_prevista"],
            mode="lines+markers",
            name=f"Venda Prevista - {linha}",
            yaxis="y1"
        ))

    fig.update_layout(
        title="Previsão de Vendas e Estoque Recomendado por Linha OTB (Próximos 6 meses)",
        xaxis_title="Ano-Mês",
        yaxis_title="Quantidade",
        barmode="group",
        legend_title_text="Legenda",
        height=600,
        xaxis=dict(
            tickformat="%Y-%m",
            dtick="M1",
            tickangle=45,
        ),
        yaxis=dict(
            title="Quantidade",
            side="left",
            showgrid=True,
            zeroline=True,
        )
    )

    st.plotly_chart(fig, use_container_width=True)
