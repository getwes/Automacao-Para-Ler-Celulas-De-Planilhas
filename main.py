#instalando as bibliotecas
# pip install openpyxl python-docx

from openpyxl import load_workbook
from docx import Document
from datetime import datetime

planilha_fornecedores = load_workbook('./fornecedores.xlsx')
pagina_fornecedores = planilha_fornecedores['Sheet1']

for linha in pagina_fornecedores.iter_rows(min_row=2,values_only=True):
    nome_empresa, endereco, cidade, estado, cep, telefone, email, setor = linha #unpacking
  
    arquivo_word = Document()
    arquivo_word.add_heading('contrato de prestação de serviço',0)