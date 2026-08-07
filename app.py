import streamlit as st

# =====================================================================
# FUNÇÕES DE SUPORTE (Garantem que nenhuma variável fique faltando)
# =====================================================================
def carregar_anuncios_por_bairro(bairro, cidade):
    """Simula ou busca a base de comparativos da região."""
    # Retorna uma lista de exemplo segura para evitar tela branca caso o banco não responda
    return [
        {"preco": 250000.0, "area_construida": 100.0},
        {"preco": 320000.0, "area_construida": 130.0}
    ]

def calcular_faixa_bairro_investidor(imovel_alvo, lista_anuncios_bairro):
    """Calcula o piso, teto e média do metro quadrado restrito ao bairro."""
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

# =====================================================================
# RENDERIZAÇÃO DA ABA DE CARDS DO INVESTIDOR
# =====================================================================
def renderizar_aba_cards_investidor(investidor_selecionado, lista_imoveis_do_investidor):
    try:
        st.markdown(f"### 🎯 Vitrine Exclusiva: {investidor_selecionado}")
        
        if not lista_imoveis_do_investidor:
            st.info("Nenhum imóvel cadastrado para este investidor.")
            return

        col1, col2 = st.columns(2)
        
        for i, imovel in enumerate(lista_imoveis_do_investidor):
            bairro = imovel.get('bairro', 'Centro')
            cidade = imovel.get('cidade', 'São José do Rio Preto - SP')
            
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
                        
    except Exception as e:
        st.error(f"Ocorreu um erro crítico ao carregar os cards: {e}")
