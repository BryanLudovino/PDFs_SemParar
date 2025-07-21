import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, Table as PlatypusTable
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import sem_parar


def resource_path(relative_path):
    """ Obtém o caminho absoluto para o recurso, funciona para desenvolvimento e para PyInstaller. """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def gerar_recibos(modelo, lista_rotas, caminho_pasta, progress_callback=None):
    """
    Gera os recibos em PDF para uma lista de rotas.
    """
    os.makedirs(caminho_pasta, exist_ok=True)

    # Remove "Fora de Rota" da lista de rotas, se necessário
    lista_rotas = [r for r in lista_rotas if r != "Fora de Rota"]

    # Remove linhas indesejadas da coluna Descrição
    descricoes_remover = [
        'E-BOOK',
        'MENSALIDADE',
        'PARCERIA SERVIÇOS VEICULARES',
        'SKEELO AUDIOBOOKS'
    ]
    modelo = modelo[~modelo['Descrição'].isin(descricoes_remover)]

    # Estilos
    styles = getSampleStyleSheet()
    centered_title = ParagraphStyle(name='CenteredTitle', parent=styles['Title'], alignment=TA_CENTER, fontSize=15, spaceAfter=10, leftIndent=-30)
    left_normal = ParagraphStyle(name='LeftNormal', parent=styles['Normal'], alignment=TA_LEFT, fontSize=12.5)
    left_normal_esquerda = ParagraphStyle(name='LeftNormalEsquerda', parent=styles['Normal'], alignment=TA_LEFT, fontSize=12.3, leftIndent=-4)
    footer_font10 = ParagraphStyle(name='FooterFont10', parent=styles['Normal'], alignment=TA_LEFT, fontSize=8)
    footer_font7 = ParagraphStyle(name='FooterFont7', parent=styles['Normal'], alignment=TA_LEFT, fontSize=8)

    total_rotas = len(lista_rotas)
    for i, id_rota in enumerate(lista_rotas):
        # Filtra as linhas da rota
        dados_filtrados = modelo[modelo['Id da Rota'] == id_rota]
        if dados_filtrados.empty:
            continue  # Não gera PDF se não houver dados
        dados = [
            ["Data", "Veículo", "Tipo", "Descrição", "Permanência", "Valor (R$)"]
        ]
        for _, row in dados_filtrados.iterrows():
            dados.append([
                row['Datax'],
                row['Placa Veículo'],
                row['Tipo'],
                row['Descrição'],
                "",
                row['Valor(R$)']
            ])
        
        # PDF Setup
        nome_arquivo = f"{id_rota}.pdf"
        caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        doc = SimpleDocTemplate(caminho_arquivo, pagesize=A4, rightMargin=65, leftMargin=35, topMargin=24, bottomMargin=30)
        elements = []

        # Cabeçalho
        logo_path = resource_path("abc.png")
        try:
            img = Image(logo_path, width=60, height=60)
            header_table = PlatypusTable([
                [img, Paragraph("Relatório de Lançamentos", centered_title), ""]
            ], colWidths=[100, 340, 60])
        except:
            header_table = PlatypusTable([
                ["", Paragraph("Relatório de Lançamentos", centered_title), ""]
            ], colWidths=[100, 340, 60])
        header_table.setStyle(TableStyle([
            ("ALIGN", (0,0), (0,0), "LEFT"),
            ("ALIGN", (1,0), (1,0), "CENTER"),
            ("ALIGN", (2,0), (2,0), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (0,0), 19),
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 13))

        # Informações da empresa
        elements.append(Paragraph("Nome: PRA LOG TRANSPORTES E SERVIÇOS LTDA ME", left_normal_esquerda))
        elements.append(Spacer(1, 7))
        elements.append(Paragraph("CPF/CNPJ: 31.882.636/0001-38", left_normal_esquerda))
        elements.append(Spacer(1, 18))

        # Tabela de dados
        tabela = Table(dados, colWidths=[75, 46, 59, 200, 55, 55])
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 7.1),
            # Bordas externas grossas para toda a tabela
            ('BOX', (0, 0), (-1, -1), 1, colors.grey),
            # Grade interna grossa para o corpo
            ('INNERGRID', (0, 1), (-1, -1), 1, colors.grey),
            # Grade interna fina só no cabeçalho
            ('INNERGRID', (0, 0), (-1, 0), 0.3, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 1.7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0.6),
        ]))
        elements.append(tabela)
        elements.append(Spacer(1, 17))

        # Rodapé
        elements.append(Paragraph("Atenciosamente,", footer_font10))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("CGMP - CENTRO DE GESTÃO DE MEIOS DE PAGAMENTO LTDA.", footer_font10))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("AVENIDA DRA RUTH CARDOSO, 7221 | PINHEIROS - SP | CEP: 05425-902", footer_font10))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("Central de Atendimento | 4002 1552 (Capitais e Regiões Metropolitanas) | 0800 015 0252 (Demais localidades)", footer_font7))

        # Gera o PDF
        doc.build(elements)

        if progress_callback:
            percent = int(((i + 1) / total_rotas) * 100)
            progress_callback(percent, f"Gerando PDF {i + 1}/{total_rotas}...")

if __name__ == '__main__':
    # Caminho da pasta de saída
    caminho_pasta_saida = r"C:\Users\Bryan Souza\Downloads\Nova pasta"

    # Chama a função para obter modelo e lista_rotas
    modelo_df, lista_rotas_df = sem_parar.fluxo_principal(
        "Maio 2025.xlsx",
        "Maio 2025 - vickos.xlsx",
        "Análise Relatório Sem Parar.xlsx"
    )
    # Remove o print da coluna
    # print(modelo.columns)
    gerar_recibos(modelo_df, lista_rotas_df, caminho_pasta_saida)
