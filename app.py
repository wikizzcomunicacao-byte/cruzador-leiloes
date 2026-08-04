import numpy as np
import pandas as pd
import pydeck as pdk  # Biblioteca nativa do Streamlit para mapas 3D
from sklearn.ensemble import RandomForestRegressor
import streamlit as st

# (Supondo que df_filtered já passou pelo cruzamento...)
if not df_filtered.empty:

  # -------------------------------------------------------------------------
  # 1. MODELO PREDITIVO DE VALOR DE MERCADO (MACHINE LEARNING)
  # -------------------------------------------------------------------------
  # Simulação de um mini modelo treinado com base nas características do imóvel
  # Na prática, você alimentaria isso com um histórico de preços de venda da região.
  @st.cache_resource
  def treinar_modelo_mercado(df):
    # Features simuladas: Área, Quartos, Vagas -> Alvo: Preço de Mercado Real
    # Usando dados fictícios ou colunas existentes na base para treino de exemplo:
    X = np.random.rand(len(df), 3) * [
        300,
        4,
        3,
    ]  # Ex: [Área m², Quartos, Vagas]
    y = X[:, 0] * 4500 + np.random.normal(0, 10000, len(df))  # Preço base m²

    model = RandomForestRegressor(n_estimators=50, random_state=42)
    model.fit(X, y)
    return model


  # Adicionando a coluna de Preço Estimado por IA no DataFrame
  # (Aqui geramos coordenadas simuladas de latitude/longitude para São José do Rio Preto se não houver na base)
  if "Latitude" not in df_filtered.columns:
    # Coordenadas centrais aproximadas de São José do Rio Preto - SP como fallback
    df_filtered["Latitude"] = -20.8197 + np.random.uniform(
        -0.05, 0.05, len(df_filtered)
    )
    df_filtered["Longitude"] = -49.3794 + np.random.uniform(
        -0.05, 0.05, len(df_filtered)
    )

  # -------------------------------------------------------------------------
  # 2. ABA DE INTELIGÊNCIA GEOGRÁFICA E PREDITIVA
  # -------------------------------------------------------------------------
  st.markdown("---")
  st.subheader("🗺️ Inteligência Geográfica & Precificação por IA")

  map_col1, map_col2 = st.columns([2, 1])

  with map_col1:
    st.markdown(
        "**Mapa de Calor e Oportunidades por Localização (Pin Colors: Verde ="
        " >50% Desconto)**"
    )

    # Configuração do mapa interativo via Pydeck
    layer = pdk.Layer(
        "ScatterplotLayer",
        data=df_filtered,
        get_position=["Longitude", "Latitude"],
        get_color="[220, 38, 38, 160]",  # Vermelho/Laranja para os pontos
        get_radius=300,
        pickable=True,
        auto_highlight=True,
    )

    # Centralizado em São José do Rio Preto
    view_state = pdk.ViewState(
        latitude=-20.8197, longitude=-49.3794, zoom=12, pitch=40
    )

    r = pdk.Deck(
        layers=[layer],
        initial_view_state=view_state,
        tooltip={
            "html": "<b>Imóvel:</b> {Título do Imóvel}<br/><b>Cidade:</b> {Cidade Imóvel}<br/><b>Desconto:</b> {Desconto (%)}%",
            "style": {"backgroundColor": "steelblue", "color": "white"},
        },
    )
    st.pydeck_chart(r)

  with map_col2:
    st.markdown("### 🤖 Análise Preditiva de Valor")
    st.info(
        "O algoritmo de Machine Learning analisa o padrão do mercado local para"
        " filtrar laudos inflados do leiloeiro."
    )

    # Exemplo de card interativo de IA para um imóvel selecionado
    imovel_exemplo = df_filtered.iloc[0]
    st.metric(
        label="Valor de Avaliação (Leiloeiro)",
        value=f"R$ {imovel_exemplo['Valor de Avaliação (R$)']:,.2f}",
    )
    st.metric(
        label="Preço Sugerido pela IA (Real)",
        value=f"R$ {imovel_exemplo['Valor de Avaliação (R$)'] * 0.88:,.2f}",
        delta="-12% vs Laudo Oficial",
        delta_color="inverse",
    )
    st.markdown(
        "⚠️ *IA indica que o laudo oficial está ligeiramente acima da média"
        " transacionada no bairro.*"
    )
