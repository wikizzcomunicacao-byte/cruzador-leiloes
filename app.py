import streamlit as st

# Configuração inicial da página
st.set_page_config(page_title="Inteligência Imobiliária - Cards", layout="wide")

def analisar_diferenciais(imovel):
    """
    Atribui um fator de ajuste com base em diferenciais do imóvel
    (ex: área gourmet, número de dormitórios, estado de conservação).
    """
    fator = 1.0
    if imovel.get("tem_area_gourmet", False):
        fator += 0.08
    if imovel.get("quartos", 3) > 3:
        fator += 0.05
    if imovel.get("padrao_acabamento") == "Alto":
        fator += 0.12
    elif imovel.get("padrao_acabamento") == "Baixo":
        fator -= 0.10
    return fator

def calcular_faixa_bairro(imovel_alvo, lista_anuncios_bairro):
    """
    Calcula a faixa de preço (R$ X a R$ Y) do bairro, a média do m²
    e o valor de mercado sugerido com base nos comparativos da região.
    """
    if not lista_anuncios_bairro:
        return {
            "faixa_bairro": "R$ 0,00 a R$ 0,00",
            "media_m2_bairro": "R$ 0,00/m²",
            "valor_mercado_sugerido": imovel_alvo["area_construida"] * 2500
        }

    precos = [a['preco'] for a in lista_anuncios_bairro]
    valores_m2 = [a['preco'] / a['area_construida'] for a in lista_anuncios_bairro]
    
    faixa_min = min(precos)
    faixa_max = max(precos)
    media_m2 = sum(valores_m2) / len(valores_m2)
    
    fator_ajuste = analisar_diferenciais(imovel_alvo)
    valor_estimado_mercado = (imovel_alvo['area_construida'] * media_m2) * fator_ajuste
    
    return {
        "faixa_bairro": f"R$ {faixa_min:,.2f} a R$ {faixa_max:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        "media_m2_bairro": f"R$ {media_m2:,.2f}/m²".replace(",", "X").replace(".", ",").replace("X", "."),
        "valor_mercado_sugerido": valor_estimado_mercado
    }

# ==========================================
# DADOS DE EXEMPLO (Simulando o Banco / API)
# ==========================================
anuncios_bairro_exemplo = [
    {"preco": 260000.0, "area_construida": 110.0},
    {"preco": 290000.0, "area_construida": 125.0},
    {"preco": 220000.0, "area_construida": 90.0},
    {"preco": 310000.0, "area_construida": 140.0}
]

imoveis_vitrine = [
    {
        "titulo": "Casa Residencial - Bairro Solo Sagrado",
        "cidade": "São José do Rio Preto - SP",
        "preco_oferta": 210000.0,
        "avaliacao_perito": 280000.0,
        "area_construida": 120.0,
        "area_terreno": 250.0,
        "tem_area_gourmet": True,
        "quartos": 3,
        "padrao_acabamento": "Médio"
    },
    {
        "titulo": "Residencial Gaivota I (Casa em Condomínio)",
        "cidade": "São José do Rio Preto - SP",
        "preco_oferta": 921005.75,
        "avaliacao_perito": 1535009.60,
        "area_construida": 247.0,
        "area_terreno": 400.0,
        "tem_area_gourmet": True,
        "quartos": 4,
        "padrao_acabamento": "Alto"
    }
]

# ==========================================
# INTERFACE DA APLICAÇÃO (STREAMLIT)
# ==========================================
st.title("🎯 Vitrine Exclusiva: Inteligência Imobiliária por Bairro")
st.markdown("Análise automatizada de faixa de mercado, metro quadrado e margem de desconto.")

col1, col2 = st.columns(2)

for i, imovel in enumerate(imoveis_vitrine):
    # Executa o cálculo da faixa de mercado do bairro
    dados_mercado = calcular_faixa_bairro(imovel, anuncios_bairro_exemplo)
    
    # Cálculo do Desconto e Múltiplo
    desconto_pct = ((imovel["avaliacao_perito"] - imovel["preco_oferta"]) / imovel["avaliacao_perito"]) * 100
    multiplo_m2 = imovel["preco_oferta"] / imovel["area_construida"]
    
    # Seleção de coluna para os cards
    col = col1 if i % 2 == 0 else col2
    
    with col:
        with st.container(border=True):
            st.markdown(f"### 🔥 {desconto_pct:.2f}% OFF | 🏠 Imóvel")
            st.markdown(f"**{imovel['titulo']}**")
            st.text(f"📍 {imovel['cidade']}")
            
            st.divider()
            
            # Dados Financeiros Principais
            st.metric(
                label="Preço de Oferta (Lance)", 
                value=f"R$ {imovel['preco_oferta']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            
            st.markdown(f"**Avaliação Original (Perito):** R$ {imovel['avaliacao_perito']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            
            st.divider()
            
            # Bloco de Inteligência por Bairro (A novidade implementada)
            st.markdown("📊 **Parâmetros de Mercado do Bairro:**")
            st.markdown(f"- **Faixa do Bairro:** {dados_mercado['faixa_bairro']}")
            st.markdown(f"- **Média M² da Região:** {dados_mercado['media_m2_bairro']}")
            st.markdown(f"- **Múltiplo deste Imóvel:** R$ {multiplo_m2:,.2f}/m²".replace(",", "X").replace(".", ",").replace("X", "."))
            
            st.divider()
            
            # Botões de Ação do Card
            b_col1, b_col2 = st.columns(2)
            with b_col1:
                st.button("Ver Anúncio", key=f"anuncio_{i}", use_container_width=True)
            with b_col2:
                st.button("WhatsApp", key=f"zap_{i}", use_container_width=True)
