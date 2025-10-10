import os
import re
import pdfplumber
from openpyxl import Workbook

# Caminhos das pastas
pdf_folder = r"C:\Users\dsadm\Documents\pedro\samples"
output_excel = r"C:\Users\dsadm\Documents\pedro\PdfToExcel\curriculos_extraidos.xlsx"

# Criar planilha Excel
wb = Workbook()
ws = wb.active
ws.title = "Currículos"
ws.append(["Nome", "Telefone", "E-mail", "Objetivo", "Arquivo PDF"])

# Regex aprimorado
regex_telefone = r'\(?\d{2}\)?\s?\d{4,5}[-.\s]?\d{4}'
regex_email = r'[\w\.-]+@[\w\.-]+\.\w+'

# Função de extração
def extrair_dados(caminho_pdf):
    texto = ""
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"

    # Limpar múltiplos espaços
    texto = re.sub(r'\s+', ' ', texto).strip()

    # Extrair nome (primeira linha em maiúsculas)
    nome_match = re.search(r'^[A-ZÁÉÍÓÚÃÕÂÊÎÔÛÇ ]{3,}', texto)
    nome = nome_match.group(0).strip() if nome_match else "Não encontrado"

    telefone = re.search(regex_telefone, texto)
    email = re.search(regex_email, texto)
    telefone = telefone.group(0) if telefone else "Não encontrado"
    email = email.group(0) if email else "Não encontrado"

    # Extrair objetivo entre "OBJETIVO" e a próxima seção
    objetivo_match = re.search(r'OBJETIVOS?\s*(.*?)\s*(IDIOMAS|FORMAÇÃO|EXPERIÊNCIAS|EXPERIÊNCIA|PROJETOS|$)', texto, re.IGNORECASE)
    objetivo = objetivo_match.group(1).strip() if objetivo_match else "Não encontrado"

    return nome, telefone, email, objetivo

# Loop pelos PDFs
for arquivo in os.listdir(pdf_folder):
    if arquivo.lower().endswith(".pdf"):
        caminho_pdf = os.path.join(pdf_folder, arquivo)
        nome, telefone, email, objetivo = extrair_dados(caminho_pdf)
        ws.append([nome, telefone, email, objetivo, arquivo])
        print(f"✅ Extraído: {arquivo}")

# Salvar planilha
wb.save(output_excel)
print(f"\n📁 Planilha criada em: {output_excel}")
