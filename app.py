from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def gerar_informativo_lcr(nome_arquivo="informativo_lcr_completo.pdf", lista_imoveis=None):
    # Configuração da página (tamanho Carta, margens reduzidas para otimizar o layout)
    doc = SimpleDocTemplate(
        nome_arquivo,
        pagesize=letter,
        rightMargin=30, leftMargin=30,
        topMargin=30, bottomMargin=30
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Estilos personalizados alinhados ao padrão visual LCR
    estilo_titulo = ParagraphStyle(
        'TituloLCR',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13,
        textColor=colors.HexColor('#2c3e50'),
        spaceAfter=4
    )
    
    estilo_texto = ParagraphStyle(
        'TextoLCR',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        textColor=colors.HexColor('#333333'),
        leading=13
    )
    
    estilo_destaque = ParagraphStyle(
        'DestaqueLCR',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        textColor=colors.HexColor('#1a252f'),
        leading=13
    )

    if not lista_imoveis:
        lista_imoveis = []

    for idx, imovel in enumerate(lista_imoveis):
        # 1. Cabeçalho Fixo com o Logotipo / Nome da Empresa
        tabela_topo = Table([
            [Paragraph("<b>LCR</b>", estilo_titulo), Paragraph("<b>LOURENÇO COLOMBO E ROZANI</b><br/><font size=7 color='#666666'>ADVOCACIA E LEILÕES IMOBILIÁRIOS</font>", estilo_texto)]
        ], colWidths=[50, 490])
        tabela_topo.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LINEAFTER', (0,0), (0,0), 1.2, colors.HexColor('#2c3e50')),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        story.append(tabela_topo)
        story.append(Spacer(1, 8))
        
        # 2. Espaço Reservado para a Foto do Imóvel (Simulando o bloco da foto)
        # Nota: Aqui você pode substituir por um objeto Image do ReportLab apontando para o arquivo da foto.
        tabela_foto = Table([[Paragraph(f"<font color='#888888'><b>[FOTO DO IMÓVEL: {imovel.get('titulo', '')}]</b></font>", ParagraphStyle('Center', alignment=1))]], colWidths=[540], rowHeights=[140])
        tabela_foto.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#ecf0f1')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#bdc3c7')),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#bdc3c7')),
        ]))
        story.append(tabela_foto)
        story.append(Spacer(1, 10))
        
        # 3. Título do Imóvel (Número e Nome)
        story.append(Paragraph(f"<b>{imovel.get('numero', '01')}. {imovel.get('titulo', 'IMÓVEL')}</b>", estilo_titulo))
        story.append(Spacer(1, 4))
        
        # 4. Dados Estruturados em Duas Colunas
        dados_tabela = [
            [Paragraph(f"<b>Matrícula:</b> {imovel.get('matricula', '-')}", estilo_texto), Paragraph(f"<b>Área do Terreno:</b> {imovel.get('terreno', '-')}", estilo_texto)],
            [Paragraph(f"<b>Quartos:</b> {imovel.get('quartos', '-')}", estilo_texto), Paragraph(f"<b>Área Construída:</b> {imovel.get('construida', '-')}", estilo_texto)],
            [Paragraph(f"<b>Vagas de Garagem:</b> {imovel.get('vagas', '-')}", estilo_texto), Paragraph(f"<b>Valor de Avaliação:</b> {imovel.get('avaliacao', '-')}", estilo_destaque)],
            [Paragraph(f"<b>Instituição:</b> {imovel.get('instituicao', '-')}", estilo_texto), Paragraph(f"<b>Lance Mínimo:</b> {imovel.get('lance_min', '-')}", estilo_destaque)],
            [Paragraph(f"<b>Modalidade:</b> {imovel.get('modalidade', '-')}", estilo_texto), Paragraph(f"<b>Pagamento:</b> {imovel.get('pagamento', 'À vista')}", estilo_texto)],
        ]
        
        tabela_dados = Table(dados_tabela, colWidths=[270, 270])
        tabela_dados.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(tabela_dados)
        story.append(Spacer(1, 6))
        
        # 5. Informações Complementares (Datas, Observações, Endereço e Link)
        story.append(Paragraph(f"<b>Data do Leilão:</b> {imovel.get('data', '-')}", estilo_texto))
        if imovel.get('obs'):
            story.append(Paragraph(f"<b>OBS:</b> {imovel.get('obs')}", estilo_texto))
        story.append(Paragraph(f"<b>Endereço:</b> {imovel.get('endereco', '-')}", estilo_texto))
        story.append(Paragraph(f"<b>Acesse o Link:</b> <font color='blue'><u>{imovel.get('link', 'Clique aqui')}</u></font>", estilo_texto))
        
        # Quebra de página automática para o próximo imóvel (mantendo o padrão de 1 imóvel por página)
        if idx < len(lista_imoveis) - 1:
            story.append(PageBreak())

    doc.build(story)
    print("Informativo gerado com sucesso!")

# Exemplo de massa de dados aplicando alguns dos imóveis que selecionamos:
if __name__ == "__main__":
    imoveis_exemplo = [
        {
            "numero": "02",
            "titulo": "ALPHAVILLE MIRASSOL",
            "matricula": "53.976",
            "quartos": "03",
            "vagas": "Não Informado",
            "terreno": "306,25 m²",
            "construida": "248,80 m²",
            "avaliacao": "R$ 1.160.000,00",
            "lance_min": "R$ 808.723,45",
            "instituicao": "Caixa",
            "modalidade": "SFI - Edital Único",
            "data": "03/09/2025 e 10/09/2025",
            "pagamento": "À vista",
            "obs": "Permite utilização de FGTS",
            "endereco": "R. Alessandra Bersi, 285, LT 26 - QD 16, Alphaville Mirassol, Mirassol/SP, CEP: 15134-180",
            "link": "https://www.caixa.gov.br"
        },
        {
            "numero": "03",
            "titulo": "PARQUE DA LIBERDADE VI",
            "matricula": "142.754",
            "quartos": "02",
            "vagas": "02",
            "terreno": "100,00 m²",
            "construida": "51,05 m²",
            "avaliacao": "R$ 230.000,00",
            "lance_min": "R$ 138.000,00",
            "instituicao": "Caixa",
            "modalidade": "SFI - Edital Único",
            "data": "13/10/2025 e 20/10/2025",
            "pagamento": "À vista",
            "obs": "Permite utilização de FGTS",
            "endereco": "Avenida José da Silva Sé, 505, Casa 47 - Qd. A, Parque da Liberdade VI, São José do Rio Preto/SP, CEP: 15056-750",
            "link": "https://www.caixa.gov.br"
        }
    ]
    
    gerar_informativo_lcr("informativo_imoveis_lcr.pdf", imoveis_exemplo)
