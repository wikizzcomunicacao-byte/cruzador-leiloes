import streamlit as st

# Configuração da página
st.set_page_config(page_title="Cruzamento de Dados de Leilões e Investidores", layout="wide")

# ==========================================
# 1. FUNÇÕES DE SUPORTE E INTELIGÊNCIA DE BAIRRO
# ==========================================
def carregar_anuncios_por_bairro(bairro, cidade):
    """Simula a varredura de anúncios de comparativos na mesma região/bairro."""
    # Em produção, aqui você faz a consulta na sua base de dados ou API
    if "Solo Sagrado" in bairro:
        return [
            {"preco": 160000.0, "area_construida": 75.0},
            {"preco": 220000.0, "area_construida": 100.0},
            {"preco": 290000.0, "area_construida": 130.0}
        ]
    elif "Gaivota" in bairro:
        return [
            {"preco": 850000.0, "area_construida": 220.0},
            {"preco": 1100000.0, "area_construida": 260.0},
            {"preco": 1500000.0, "area_construida": 300.0}
        ]
    else:
        return [
            {"preco": 200000.0, "area_construida": 90.0},
            {"preco": 350000.0, "area_construida": 150.0}
        ]

def calcular_faixa_bairro_investidor(imovel_alvo, lista_anuncios_bairro):
    """Calcula o piso, teto e média do metro quadrado restrito ao bairro do imóvel."""
    if not lista_anuncios_bairro:
        return {
            "faixa_bairro": "R$ 0,00 a R$ 0,00",
            "media_m2_bairro": "R$ 0,00/m²",
        }

    precos = [a['preco'] for a in lista_anuncios_bairro]
    valores_m2 = [a['preco'] / a['area_construida'] for a in lista_anuncios_bairro]
    
    faixa_min = min(precos)
    faixa_max = max(precos)
    media_m2 = sum(valores_m2) / len(valores_m2)
    
    return {
        "faixa_bairro": f"R$ {faixa_min:,.2f} a R$ {faixa_max:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        "media_m2_bairro": f"R$ {media_m2:,.2f}/m²".replace(",", "X").replace(".", ",").replace("X", ".")
    }

# ==========================================
# 2. BASE DE DADOS SIMULADA (Seu Projeto Completo)
# ==========================================
if "imoveis" not in st.session_state:
    st.session_state["imoveis"] = [
        {
            "titulo": "Casa Residencial - Bairro Solo Sagrado",
            "bairro": "Solo Sagrado",
            "cidade": "São José do Rio Preto - SP",
            "investidor": "Assad Gabriel Assad Neto",
            "preco_oferta": 210000.0,
            "avaliacao_perito": 280000.0,
            "area_construida": 120.0
        },
        {
            "titulo": "Residencial Gaivota I (Casa em Condomínio)",
            "bairro": "Residencial Gaivota I",
            "cidade": "São José do Rio Preto - SP",
            "investidor": "Assad Gabriel Assad Neto",
            "preco_oferta": 921005.75,
            "avaliacao_perito": 1535009.60,
            "area_construida": 247.0
        }
    ]

# ==========================================
# 3. INTERFACE PRINCIPAL E ABAS
# ==========================================
st.title("🎯 Sistema de Cruzamento de Leilões e Imóveis")

aba1, aba2 = st.tabs(["📊 Visão Geral / Tabela", "📑 Cards por Investidor (PDF / WhatsApp)"])

with aba1:
    st.subheader("Base Geral de Imóveis Cadastrados")
    st.dataframe(st.session_state["imoveis"], use_container_width=True)

with aba2:
    st.subheader("Vitrine Dinâmica por Investidor")
    
    # Filtro de Investidor
    investidores_disponiveis = list(set([im["investidor"] for im in st.session_state["imoveis"]]))
    investidor_selecionado = st.selectbox("Selecione o Investidor:", investidores_disponiveis)
    
    # Filtra os imóveis do investidor escolhido
    lista_imoveis_do_investidor = [im for im in st.session_state["imoveis"] if im["investidor"] == investidor_selecionado]
    
    if not lista_imoveis_do_investidor:
        st.info("Nenhum imóvel encontrado para este investidor.")
    else:
        st.markdown(f"### 🎯 Vitrine Exclusiva: {investidor_selecionado}")
        
        col1, col2 = st.columns(2)
        
        for i, imovel in enumerate(lista_imoveis_do_investidor):
            bairro = imovel.get('bairro', 'Centro')
            cidade = imovel.get('cidade', 'São José do Rio Preto - SP')
            
            # Executa a busca e cálculo por bairro
            anuncios_do_bairro = carregar_anuncios_por_bairro(bairro, cidade)
            dados_mercado = calcular_faixa_bairro_investidor(imovel, anuncios_do_bairro)
            
            avaliacao = imovel.get("avaliacao_perito", 1.0)
            oferta = imovel.get("preco_oferta", 0.0)
            area = imovel.get("area_construida", 1.0)
            
            desconto_pct = ((avaliacao - oferta) / avaliacao) * 100 if avaliacao > 0 else 0
            multiplo_m2 = oferta / area if area > 0 else 0
            
            col = col1 if i % 2 == 0 else col2
            
            with col:
                with st.container(border=True):
                    st.markdown(f"### 🔥 {desconto_pct:.2f}% OFF | 🏠 Imóvel")
                    st.markdown(f"**{imovel.get('titulo', 'Imóvel')}**")
                    st.text(f"📍 {bairro} - {cidade}")
                    
                    st.divider()
                    
                    st.metric(
                        label="Preço de Oferta (Lance)", 
                        value=f"R$ {oferta:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    )
                    st.markdown(f"**Avaliação Original:** R$ {avaliacao:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    
                    st.divider()
                    
                    # Bloco Integrado de Inteligência de Bairro (R$ X a R$ Y e Média M²)
                    st.markdown("📊 **Inteligência de Mercado (Bairro):**")
                    st.markdown(f"- **Faixa da Região:** {dados_mercado['faixa_bairro']}")
                    st.markdown(f"- **Média M² do Bairro:** {dados_mercado['media_m2_bairro']}")
                    st.markdown(f"- **Múltiplo deste Imóvel:** R$ {multiplo_m2:,.2f}/m²".replace(",", "X").replace(".", ",").replace("X", "."))
                    
                    st.divider()
                    
                    b_col1, b_col2 = st.columns(2)
                    with b_col1:
                        st.button("Ver Anúncio", key=f"inv_anuncio_{i}", use_container_width=True)
                    with b_col2:
                        st.button("WhatsApp", key=f"inv_zap_{i}", use_container_width=True)
