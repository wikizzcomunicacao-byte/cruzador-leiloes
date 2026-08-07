import streamlit as st

st.set_page_config(page_title="Cruzamento de Dados de Investidores", layout="wide")

st.title("🎯 Vitrine Exclusiva: Assad Gabriel Assad Neto")
st.markdown("Análise automatizada de faixa de mercado, metro quadrado e margem de desconto.")

# Dados simulados para testar a tela
imoveis_exemplo = [
    {
        "titulo": "Casa Residencial - Bairro Solo Sagrado",
        "bairro": "Solo Sagrado",
        "cidade": "São José do Rio Preto - SP",
        "preco_oferta": 210000.0,
        "avaliacao_perito": 280000.0,
        "area_construida": 120.0
    },
    {
        "titulo": "Residencial Gaivota I (Casa em Condomínio)",
        "bairro": "Residencial Gaivota I",
        "cidade": "São José do Rio Preto - SP",
        "preco_oferta": 921005.75,
        "avaliacao_perito": 1535009.60,
        "area_construida": 247.0
    }
]

col1, col2 = st.columns(2)

for i, imovel in enumerate(imoveis_exemplo):
    avaliacao = imovel["avaliacao_perito"]
    oferta = imovel["preco_oferta"]
    area = imovel["area_construida"]
    
    desconto_pct = ((avaliacao - oferta) / avaliacao) * 100
    multiplo_m2 = oferta / area
    
    # Simulação da faixa de bairro baseada na nossa conversa
    faixa_bairro = "R$ 150.000,00 a R$ 300.000,00" if i == 0 else "R$ 850.000,00 a R$ 1.600.000,00"
    media_m2 = "R$ 2.050,00/m²" if i == 0 else "R$ 3.730,00/m²"
    
    col = col1 if i % 2 == 0 else col2
    
    with col:
        with st.container(border=True):
            st.markdown(f"### 🔥 {desconto_pct:.2f}% OFF | 🏠 Imóvel")
            st.markdown(f"**{imovel['titulo']}**")
            st.text(f"📍 {imovel['bairro']} - {imovel['cidade']}")
            
            st.divider()
            
            st.metric(
                label="Preço de Oferta (Lance)", 
                value=f"R$ {oferta:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            st.markdown(f"**Avaliação Original:** R$ {avaliacao:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            
            st.divider()
            
            st.markdown("📊 **Inteligência de Mercado (Bairro):**")
            st.markdown(f"- **Faixa da Região:** {faixa_bairro}")
            st.markdown(f"- **Média M² do Bairro:** {media_m2}")
            st.markdown(f"- **Múltiplo deste Imóvel:** R$ {multiplo_m2:,.2f}/m²".replace(",", "X").replace(".", ",").replace("X", "."))
            
            st.divider()
            
            b_col1, b_col2 = st.columns(2)
            with b_col1:
                st.button("Ver Anúncio", key=f"anuncio_{i}", use_container_width=True)
            with b_col2:
                st.button("WhatsApp", key=f"zap_{i}", use_container_width=True)
