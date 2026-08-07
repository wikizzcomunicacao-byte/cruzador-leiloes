# =====================================================================
# INSERIR NA SUA ABA: "Cards por Investidor (PDF / WhatsApp)"
# =====================================================================

def renderizar_aba_cards_investidor(investidor_selecionado, lista_imoveis_do_investidor):
    st.markdown(f"### 🎯 Vitrine Exclusiva: {investidor_selecionado}")
    
    # Divide a tela em colunas para os cards (exemplo em 2 colunas)
    col1, col2 = st.columns(2)
    
    for i, imovel in enumerate(lista_imoveis_do_investidor):
        
        # 1. FUNÇÃO DE CÁLCULO ESPECÍFICA PARA OS CARDS DO INVESTIDOR
        # (Aqui você puxa do seu banco a lista de anúncios reais do mesmo bairro do imóvel)
        anuncios_do_bairro = carregar_anuncios_por_bairro(imovel['bairro'], imovel['cidade'])
        
        dados_mercado = calcular_faixa_bairro_investidor(imovel, anuncios_do_bairro)
        
        # Cálculos de deságio e múltiplo do card
        desconto_pct = ((imovel["avaliacao_perito"] - imovel["preco_oferta"]) / imovel["avaliacao_perito"]) * 100
        multiplo_m2 = imovel["preco_oferta"] / imovel["area_construida"]
        
        # Alterna entre as colunas da interface
        col = col1 if i % 2 == 0 else col2
        
        with col:
            with st.container(border=True):
                # Cabeçalho do Card
                st.markdown(f"### 🔥 {desconto_pct:.2f}% OFF | 🏠 Imóvel")
                st.markdown(f"**{imovel['titulo']}**")
                st.text(f"📍 {imovel['bairro']} - {imovel['cidade']}")
                
                st.divider()
                
                # Valores Principais
                st.metric(
                    label="Preço de Oferta (Lance)", 
                    value=f"R$ {imovel['preco_oferta']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
                st.markdown(f"**Avaliação Original:** R$ {imovel['avaliacao_perito']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                st.divider()
                
                # 2. DADOS DE FAIXA DE BAIRRO (Aplicado somente nesta aba)
                st.markdown("📊 **Inteligência de Mercado (Bairro):**")
                st.markdown(f"- **Faixa da Região:** {dados_mercado['faixa_bairro']}")
                st.markdown(f"- **Média M² do Bairro:** {dados_mercado['media_m2_bairro']}")
                st.markdown(f"- **Múltiplo deste Imóvel:** R$ {multiplo_m2:,.2f}/m²".replace(",", "X").replace(".", ",").replace("X", "."))
                
                st.divider()
                
                # Botões de Ação
                b_col1, b_col2 = st.columns(2)
                with b_col1:
                    st.button("Ver Anúncio", key=f"inv_anuncio_{i}", use_container_width=True)
                with b_col2:
                    st.button("WhatsApp", key=f"inv_zap_{i}", use_container_width=True)


def calcular_faixa_bairro_investidor(imovel_alvo, lista_anuncios_bairro):
    """Calcula o piso, teto e média do metro quadrado restrito ao bairro do imóvel."""
    if not lista_anuncios_bairro:
        return {
            "faixa_bairro": "Dados indisponíveis",
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
