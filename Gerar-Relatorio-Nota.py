import os
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def extrair_dados_xml(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    data_emissao = root.find('.//nfe:dhEmi', ns).text

    produtos = [
        (item.find('.//nfe:cProd', ns).text,
         item.find('.//nfe:xProd', ns).text,
         int(float(item.find('.//nfe:qCom', ns).text)),
         item.find('.//nfe:uCom', ns).text)
        for item in root.findall('.//nfe:det', ns)
    ]

    return produtos, data_emissao

def gerar_pdf(produtos, imagens_dir, arquivo_pdf, data_emissao):
    pdf_doc = SimpleDocTemplate(
        arquivo_pdf,
        pagesize=letter,
        leftMargin=28.35,
        rightMargin=28.35,
        topMargin=28.35,
        bottomMargin=28.35
    )

    styles = getSampleStyleSheet()
    rodape = Paragraph(f"Data de emissão: {data_emissao}", styles['Normal'])

    dados_tabela = [('Código', 'Descrição', 'Quantidade', 'Unidade', 'Imagem')]
    
    for produto in produtos:
        codigo, descricao, quantidade, unidade = produto
        imagem_path = os.path.join(imagens_dir, f"{codigo}.jpg")
        if os.path.exists(imagem_path):
            imagem = Image(imagem_path)
            imagem.drawHeight = 50
            imagem.drawWidth = 50
        else:
            imagem = 'Sem imagem'
        dados_tabela.append((codigo, descricao, quantidade, unidade, imagem))

    tabela = Table(dados_tabela, repeatRows=1)
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elementos = [tabela, Spacer(1, 10), rodape]
    pdf_doc.build(elementos)

def escolher_arquivo(titulo, tipo_de_arquivo):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=titulo, filetypes=[tipo_de_arquivo])

def main(arquivo_xml, imagens_dir):
    produtos, data_emissao = extrair_dados_xml(arquivo_xml)
    nome_arquivo_pdf = os.path.splitext(os.path.basename(arquivo_xml))[0] + ".pdf"
    caminho_arquivo_pdf = os.path.join(os.getcwd(), nome_arquivo_pdf)
    gerar_pdf(produtos, imagens_dir, caminho_arquivo_pdf, data_emissao)

if __name__ == "__main__":
    arquivo_xml = escolher_arquivo("Selecionar arquivo XML", ("Arquivos XML", "*.xml"))
    # Define a pasta base onde o script está localizado
    base_dir = os.path.dirname(os.path.abspath(__file__))
    # Cria o caminho completo para a pasta de imagens
    imagens_dir = os.path.join(base_dir, "Imagens")
    if arquivo_xml:
        main(arquivo_xml, imagens_dir)
    else:
        print("Operação cancelada.")
