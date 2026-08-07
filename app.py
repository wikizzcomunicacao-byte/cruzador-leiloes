def renderizar_aba_cards_investidor(investidor_selecionado, lista_imoveis_do_investidor):
    st.markdown(f"### 🎯 Vitrine Exclusiva: {investidor_selecionado}")
    
    # Validação de segurança para evitar tela branca se a lista estiver vazia
    if not lista_imoveis_do_investidor:
        st.info("Nenhum imóvel encontrado para este investidor ou filtro selecionado.")
        return

    col1, col2 = st.columns(2)
    
    for i, imovel in enumerate(lista_imoveis_do_investidor):
        try:
            # Puxa os anúncios do bairro (certifique-se de que essa função existe no seu escopo)
            bairro = imovel.get('bairro', 'Geral')
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
                    st.markdown(f"**{imovel.get('titulo', 'Imóvel sem título')}**")
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
            st.error(f"Erro ao renderizar o card do imóvel {i}: {e}")
