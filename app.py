import io
import re
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------------------------------------------
# 1. CONFIGURAÇÃO DA PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Cruzador de Leilões Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------
# 2. ESTILIZAÇÃO CSS CUSTOMIZADA (SaaS Look & Feel)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    /* Oculta menus e rodapé padrão do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }
    
    /* Botão Principal */
    div.stButton > button:first-child {
        background-color: #0052CC;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        border: none;
        width: 100%;
        transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        background-color: #003D99;
        color: white;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    /* Cards de Métricas (KPIs) */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 1rem;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0052CC;
    }
    
    /* Cards Estilo Imobiliária (Aba 3) */
    .property-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .property-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 14px rgba(0,0,0,0.1);
    }
    .badge-discount {
        background-color: #DC2626;
        color: white;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: bold;
        display: inline-block;
        margin-bottom: 8px;
    }
    .badge-type {
        background-color: #E0E7FF;
        color: #3730A3;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 8px;
        margin-left: 5px;
    }
    .price-main {
        font-size: 1.4rem;
        font-weight: bold;
        color: #166534;
    }
    .price-old {
        font-size: 0.95rem;
        color: #9CA3AF;
        text-decoration: line-through;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# 3. FUNÇÕES DE TRATAMENTO E EXPORTAÇÃO
# ---------------------------------------------------------
def normalize(text):
  if not isinstance(text, str):
    return ""
  text = text.lower()
  text = re.sub(r"[áàâãä]", "a", text)
  text = re.sub(r"[éèêë]", "e", text)
  text = re.sub(r"[íìîï]", "i", text)
  text = re.sub(r"[óòôõö]", "o", text)
  text = re.sub(r"[úùûü]", "u", text)
  text = re.sub(r"[ç]", "c", text)
  text = re.sub(r"[^a-z0-9\s]", " ", text)
  return " ".join(text.split())


def parse_budget(budget_str):
  budget_str = str(budget_str)
  if "100.000" in budget_str:
    return 0, 100000
  elif "200.000" in budget_str:
    return 0, 200000
  elif "300.000" in budget_str:
    return 0, 300000
  elif "400.000" in budget_str:
    return 0, 400000
  elif "500.000" in budget_str:
    return 500000, 999999999
  else:
    return 0, 999999999


def parse_types(tipos_str):
  norm_t = normalize(tipos_str)
  res = set()
  if "apartamento" in norm_t:
    res.add("Apartamento")
  if "casa" in norm_t:
    res.add("Casa")
  if "terreno" in norm_t:
    res.add("Terreno")
  if "lote comercial" in norm_t or "comercial" in norm_t:
    res.add("Comercial")
  if "outros" in norm_t:
    res.update(["Area Rural", "Comercial", "Indefinido", "Vaga de Garagem"])
  return list(res)


def gerar_excel_profissional(df_input):
  output = io.BytesIO()
  df_export = df_input.copy()

  if "Link do Imóvel" in df_export.columns:
    df_export["Link do Imóvel"] = df_export["Link do Imóvel"].apply(
        lambda x: f'=HYPERLINK("{x}", "🔗 Ver Anúncio")'
        if pd.notnull(x) and str(x).startswith("http")
        else x
    )

  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_export.to_excel(writer, index=False, sheet_name="Oportunidades")
    ws = writer.sheets["Oportunidades"]

    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(
        start_color="0052CC", end_color="0052CC", fill_type="solid"
    )
    font_link = Font(
        name="Calibri", size=11, color="0066CC", underline="single"
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    col_names = [cell.value for cell in ws[1]]

    for cell in ws[1]:
      cell.font = font_header
      cell.fill = fill_header
      cell.alignment = align_center

    for col_idx, col_name in enumerate(col_names, start=1):
      col_letter = get_column_letter(col_idx)
      cells = ws[col_letter][1:]

      if "Preço" in str(col_name) or "Avaliação" in str(col_name):
        for c in cells:
          c.number_format = "R$ #,##0.00"
          c.alignment = align_right
      elif "Desconto" in str(col_name):
        for c in cells:
          c.number_format = '0.00"%"'
          c.alignment = align_right
      elif col_name in ["ID Investidor", "Estado Imóvel"]:
        for c in cells:
          c.alignment = align_center
      elif "Link" in str(col_name):
        for c in cells:
          c.font = font_link
          c.alignment = align_center
      else:
        for c in cells:
          c.alignment = align_left

    col_widths = {
        "ID Investidor": 14,
        "Nome do Investidor": 25,
        "Cidades Solicitadas": 25,
        "Faixa Solicitada": 22,
        "Título do Imóvel": 40,
        "Cidade Imóvel": 20,
        "Estado Imóvel": 14,
        "Tipo de Bem": 16,
        "Preço do Leilão (R$)": 22,
        "Valor de Avaliação (R$)": 24,
        "Desconto (%)": 15,
        "Endereço": 45,
        "Link do Imóvel": 18,
    }

    for col_idx, col_name in enumerate(col_names, start=1):
      col_letter = get_column_letter(col_idx)
      ws.column_dimensions[col_letter].width = col_widths.get(col_name, 20)

  return output.getvalue()


# ---------------------------------------------------------
# 4. BARRA LATERAL (CONTROLES E UPLOAD)
# ---------------------------------------------------------
with st.sidebar:
  st.title("🏢 Cruzador Pro")
  st.caption("Painel Inteligente de Oportunidades")
  st.divider()

  st.subheader("📂 Enviar Arquivos")
  file_leiloes = st.file_uploader(
      "1️⃣ Base de Leilões (.xlsx)", type=["xlsx", "xls"]
  )
  file_investidores = st.file_uploader(
      "2️⃣ Base de Investidores (.xlsx)", type=["xlsx", "xls"]
  )

  st.divider()
  executar = st.button("🚀 Processar Oportunidades", type="primary")


# ---------------------------------------------------------
# 5. CABEÇALHO PRINCIPAL E DIÁLOGO DE AJUDA
# ---------------------------------------------------------
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
  st.title("🎯 Cruzador Automático de Leilões")
  st.markdown(
      "Cruzamento inteligente entre o perfil dos investidores e as"
      " oportunidades em leilão."
  )

with col_head2:
  st.write(" ")
  with st.popover("❓ Como Usar"):
    st.markdown("""
        ### 📖 Passo a Passo Simples
        1. **Carregue as planilhas** no menu lateral à esquerda.
        2. Clique em **🚀 Processar Oportunidades**.
        3. Use os **filtros de desconto e tipo** para refinar.
        4. Navegue pelas **abas**:
           * 📊 **Visão Geral:** Gráficos e indicadores gerais.
           * 📋 **Tabela:** Dados completos e download em Excel.
           * 👤 **Por Investidor:** Ficha de cada cliente com seus imóveis.
        """)


# ---------------------------------------------------------
# 6. LÓGICA DE PROCESSAMENTO E ARMAZENAMENTO EM SESSION
# ---------------------------------------------------------
if file_leiloes and file_investidores and executar:
  with st.spinner("Analisando critérios e cruzando bases de dados..."):
    try:
      df_leiloes = pd.read_excel(file_leiloes)
      df_investidores = pd.read_excel(file_investidores)

      df_leiloes["norm_cidade"] = df_leiloes["Cidade"].apply(normalize)
      df_leiloes["norm_estado"] = df_leiloes["Estado"].apply(normalize)
      df_leiloes["norm_tipo"] = df_leiloes["Tipo de Bem"].apply(normalize)

      df_leiloes["preco_effective"] = df_leiloes.apply(
          lambda r: (
              r["2º Leilão (Preço)"]
              if pd.notnull(r["2º Leilão (Preço)"])
              and r["2º Leilão (Preço)"] > 0
              else r["1º Leilão (Preço)"]
          ),
          axis=1,
      )

      df_leiloes["desconto_%"] = np.where(
          (pd.notnull(df_leiloes["Valor de Avaliação do Leiloeiro"]))
          & (df_leiloes["Valor de Avaliação do Leiloeiro"] > 0),
          ((
              df_leiloes["Valor de Avaliação do Leiloeiro"]
              - df_leiloes["preco_effective"]
          )
          / df_leiloes["Valor de Avaliação do Leiloeiro"])
          * 100,
          0,
      )

      cols = list(df_investidores.columns)
      resultados = []

      for idx, row in df_investidores.iterrows():
        nome = str(row["Nome Completo"]).strip()
        cidades_input = str(row[cols[10]]).strip()
        tipos_input = str(row[cols[7]]).strip()
        valor_input = str(row[cols[8]]).strip()
        cons = str(row[cols[15]]).strip()

        norm_cid = normalize(cidades_input)
        norm_cons = normalize(cons)
        norm_val = normalize(valor_input)

        sub = df_leiloes.copy()

        target_cidades = []
        target_estados = []

        if any(
            w in norm_cid or w in norm_cons
            for w in [
                "rio preto",
                "sao jose do rio preto",
                "sao jose do preto",
                "sjrp",
            ]
        ):
          target_cidades.append("sao jose do rio preto")
        if "bady" in norm_cid or "bady" in norm_cons:
          target_cidades.append("bady bassitt")
        if "mirassol" in norm_cid or "mirassol" in norm_cons:
          target_cidades.append("mirassol")
        if "santo andre" in norm_cid or "santo andre" in norm_cons:
          target_cidades.append("santo andre")
        if "sao bernardo" in norm_cid or "sao bernardo" in norm_cons:
          target_cidades.append("sao bernardo do campo")
        if "sao caetano" in norm_cid or "sao caetano" in norm_cons:
          target_cidades.append("sao caetano do sul")
        if "abc" in norm_cid or "abc" in norm_cons:
          target_cidades.extend(
              ["santo andre", "sao bernardo do campo", "sao caetano do sul"]
          )
        if "sao paulo" in norm_cid or "sao paulo" in norm_cons:
          if "grande sao paulo" in norm_cid or "grande sao paulo" in norm_cons:
            target_cidades.extend([
                "sao paulo",
                "santo andre",
                "sao bernardo do campo",
                "sao caetano do sul",
            ])
          elif "estado de sp" in norm_cid or "estado de sp" in norm_cons:
            target_estados.append("sao paulo")
          else:
            target_cidades.append("sao paulo")
        if "sorocaba" in norm_cid or "sorocaba" in norm_cons:
          target_cidades.append("sorocaba")
        if "tatui" in norm_cid or "tatui" in norm_cons:
          target_cidades.append("tatui")
        if "boituva" in norm_cid or "boituva" in norm_cons:
          target_cidades.append("boituva")
        if "rio de janeiro" in norm_cid or "rio de janeiro" in norm_cons:
          target_cidades.append("rio de janeiro")
        if "goiania" in norm_cid:
          target_cidades.append("goiania")
        if "brasilia" in norm_cid:
          target_cidades.append("brasilia")
        if "anapolis" in norm_cid:
          target_cidades.append("anapolis")
        if "cidades polo" in norm_cid or "expansao" in norm_cid:
          target_estados.append("sao paulo")

        target_cidades = list(set(target_cidades))
        target_estados = list(set(target_estados))

        if target_cidades:
          sub = sub[sub["norm_cidade"].isin(target_cidades)]
        elif target_estados:
          sub = sub[sub["norm_estado"].isin(target_estados)]

        allowed_types = parse_types(tipos_input)
        if "terrenos" in norm_cons or "lotes" in norm_cons:
          if "Terreno" not in allowed_types:
            allowed_types.append("Terreno")

        if allowed_types:
          sub = sub[sub["Tipo de Bem"].isin(allowed_types)]

        min_v, max_v = parse_budget(valor_input)
        if any(
            p in norm_cons or p in norm_val
            for p in [
                "qualquer valor",
                "valores menores",
                "100 a 500",
                "todos os valores",
                "qualquer",
            ]
        ):
          min_v, max_v = 0, 999999999

        if min_v > 0 and max_v < 999999999:
          sub = sub[
              (sub["preco_effective"] >= min_v)
              & (sub["preco_effective"] <= max_v)
          ]
        elif min_v > 0:
          sub = sub[sub["preco_effective"] >= min_v]
        elif max_v < 999999999:
          sub = sub[sub["preco_effective"] <= max_v]

        sub_sorted = sub.sort_values(
            by=["desconto_%", "preco_effective"], ascending=[False, True]
        )

        for _, imovel in sub_sorted.iterrows():
          resultados.append({
              "ID Investidor": idx + 1,
              "Nome do Investidor": nome,
              "Cidades Solicitadas": cidades_input,
              "Faixa Solicitada": valor_input,
              "Título do Imóvel": imovel["Título"],
              "Cidade Imóvel": imovel["Cidade"],
              "Estado Imóvel": imovel["Estado"],
              "Tipo de Bem": imovel["Tipo de Bem"],
              "Preço do Leilão (R$)": imovel["preco_effective"],
              "Valor de Avaliação (R$)": imovel[
                  "Valor de Avaliação do Leiloeiro"
              ],
              "Desconto (%)": round(imovel["desconto_%"], 2),
              "Endereço": imovel["Endereço"],
              "Link do Imóvel": imovel["Link"],
          })

      st.session_state["df_final"] = pd.DataFrame(resultados)
      st.toast("✅ Processamento concluído com sucesso!", icon="🎉")

    except Exception as e:
      st.error(f"Erro ao processar as planilhas: {e}")


# ---------------------------------------------------------
# 7. EXIBIÇÃO DO DASHBOARD E ABAS
# ---------------------------------------------------------
if "df_final" in st.session_state and not st.session_state["df_final"].empty:
  df_base = st.session_state["df_final"]

  # BARRA DE FILTROS DINÂMICOS
  with st.expander("🔍 **Filtros de Refinamento ao Vivo**", expanded=True):
    f_col1, f_col2, f_col3 = st.columns([1, 2, 2])

    with f_col1:
      desconto_min = st.slider(
          "Desconto Mínimo (%)",
          min_value=0,
          max_value=80,
          value=0,
          step=5,
      )

    with f_col2:
      tipos_disponiveis = sorted(df_base["Tipo de Bem"].unique().tolist())
      tipos_selecionados = st.multiselect(
          "Filtrar por Tipo de Bem",
          options=tipos_disponiveis,
          default=tipos_disponiveis,
      )

    with f_col3:
      busca_texto = st.text_input(
          "Buscar (Cidade, Investidor ou Título)",
          value="",
          placeholder="Ex: Rio Preto, Simonetti, Casa...",
      )

  # Aplicar Filtros
  df_filtered = df_base[
      (df_base["Desconto (%)"] >= desconto_min)
      & (df_base["Tipo de Bem"].isin(tipos_selecionados))
  ]

  if busca_texto.strip():
    termo = normalize(busca_texto)
    df_filtered = df_filtered[
        df_filtered["Cidade Imóvel"].apply(normalize).str.contains(termo)
        | df_filtered["Nome do Investidor"].apply(normalize).str.contains(termo)
        | df_filtered["Título do Imóvel"].apply(normalize).str.contains(termo)
    ]

  # ESTRUTURA DE ABAS
  tab1, tab2, tab3 = st.tabs([
      "📊 Visão Geral & Gráficos",
      "📋 Tabela Completa & Download",
      "👤 Cards por Investidor",
  ])

  # -------------------------------------------------------
  # ABA 1: VISÃO GERAL & GRÁFICOS
  # -------------------------------------------------------
  with tab1:
    st.write(" ")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("🏢 Total de Imóveis", f"{len(df_filtered):,}")
    kpi2.metric(
        "👥 Investidores Atendidos",
        f"{df_filtered['Nome do Investidor'].nunique()}"
        if not df_filtered.empty
        else "0",
    )
    kpi3.metric(
        "🔥 Maior Desconto",
        f"{df_filtered['Desconto (%)'].max():.1f}%"
        if not df_filtered.empty
        else "0%",
    )
    avg_price = (
        df_filtered["Preço do Leilão (R$)"].mean()
        if not df_filtered.empty
        else 0
    )
    kpi4.metric("💰 Ticket Médio", f"R$ {avg_price:,.2f}")

    st.markdown("---")

    if not df_filtered.empty:
      g_col1, g_col2 = st.columns(2)

      with g_col1:
        st.subheader("🍩 Distribuição por Tipo de Imóvel")
        fig_pie = px.pie(
            df_filtered,
            names="Tipo de Bem",
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Bold,
        )
        fig_pie.update_layout(margin=dict(t=20, b=20, l=10, r=10), height=320)
        st.plotly_chart(fig_pie, use_container_width=True)

      with g_col2:
        st.subheader("📍 Top 5 Cidades com Mais Oportunidades")
        top_cidades = (
            df_filtered["Cidade Imóvel"]
            .value_counts()
            .head(5)
            .reset_index()
        )
        top_cidades.columns = ["Cidade", "Total"]
        fig_bar = px.bar(
            top_cidades,
            x="Total",
            y="Cidade",
            orientation="h",
            text="Total",
            color_discrete_sequence=["#0052CC"],
        )
        fig_bar.update_layout(
            margin=dict(t=20, b=20, l=10, r=10),
            height=320,
            yaxis={"categoryorder": "total ascending"},
        )
        st.plotly_chart(fig_bar, use_container_width=True)

  # -------------------------------------------------------
  # ABA 2: TABELA COMPLETA & DOWNLOAD EXCEL
  # -------------------------------------------------------
  with tab2:
    st.write(" ")
    if not df_filtered.empty:
      excel_bytes = gerar_excel_profissional(df_filtered)

      d_col1, d_col2 = st.columns([3, 1])
      with d_col1:
        st.subheader("📋 Tabela de Oportunidades Consolidadas")
      with d_col2:
        st.download_button(
            label="📥 Baixar Excel Formatado",
            data=excel_bytes,
            file_name="cruzamento_leiloes_investidores.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

      st.dataframe(
          df_filtered,
          use_container_width=True,
          height=500,
          column_config={
              "Link do Imóvel": st.column_config.LinkColumn(
                  "Anúncio Oficial", display_text="🔗 Ver Imóvel"
              ),
              "Preço do Leilão (R$)": st.column_config.NumberColumn(
                  "Preço Leilão", format="R$ %,.2f"
              ),
              "Valor de Avaliação (R$)": st.column_config.NumberColumn(
                  "Valor Avaliação", format="R$ %,.2f"
              ),
              "Desconto (%)": st.column_config.ProgressColumn(
                  "Desconto (%)", format="%.1f%%", min_value=0, max_value=100
              ),
          },
      )
    else:
      st.warning(
          "Nenhum imóvel encontrado com os filtros selecionados acima."
      )

  # -------------------------------------------------------
  # ABA 3: CARDS POR INVESTIDOR (VITRINE IMOBILIÁRIA)
  # -------------------------------------------------------
  with tab3:
    st.write(" ")
    if not df_filtered.empty:
      investidores_lista = sorted(
          df_filtered["Nome do Investidor"].unique().tolist()
      )
      investidor_sel = st.selectbox(
          "👤 Selecione o Investidor para visualizar sua vitrine:",
          options=investidores_lista,
      )

      df_inv = df_filtered[
          df_filtered["Nome do Investidor"] == investidor_sel
      ]

      st.markdown(f"### 🎯 Oportunidades Selecionadas para: **{investidor_sel}**")
      st.caption(
          f"Preferências solicitadas: {df_inv['Cidades Solicitadas'].iloc[0]} |"
          f" Faixa: {df_inv['Faixa Solicitada'].iloc[0]}"
      )

      st.write(" ")

      # Grid de Cards (2 por linha)
      cols_cards = st.columns(2)
      for idx, (_, row) in enumerate(df_inv.iterrows()):
        col_target = cols_cards[idx % 2]

        with col_target:
          link_url = (
              row["Link do Imóvel"]
              if str(row["Link do Imóvel"]).startswith("http")
              else "#"
          )

          st.markdown(
              f"""
                <div class="property-card">
                    <span class="badge-discount">🔥 {row['Desconto (%)']}% OFF</span>
                    <span class="badge-type">🏠 {row['Tipo de Bem']}</span>
                    <h4 style="margin-top: 8px; margin-bottom: 4px; color: #1E293B;">{row['Título do Imóvel']}</h4>
                    <p style="color: #64748B; font-size: 0.9rem; margin-bottom: 12px;">📍 {row['Cidade Imóvel']} - {row['Estado Imóvel']}</p>
                    <div style="margin-bottom: 12px;">
                        <span class="price-main">R$ {row['Preço do Leilão (R$)']:,.2f}</span><br>
                        <span class="price-old">Avaliado em: R$ {row['Valor de Avaliação (R$)']:,.2f}</span>
                    </div>
                    <p style="font-size: 0.85rem; color: #475569; margin-bottom: 15px;"><b>Endereço:</b> {row['Endereço']}</p>
                    <a href="{link_url}" target="_blank" style="text-decoration: none;">
                        <button style="background-color: #0052CC; color: white; border: none; padding: 8px 16px; border-radius: 6px; font-weight: bold; cursor: pointer; width: 100%;">
                            🔗 Abrir Anúncio Oficial
                        </button>
                    </a>
                </div>
                """,
              unsafe_allow_html=True,
          )

elif "df_final" not in st.session_state:
  st.info(
      "👈 **Para iniciar:** Faça o upload das duas planilhas no menu lateral e"
      " clique em **🚀 Processar Oportunidades**."
  )
