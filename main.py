#instalando as bibliotecas
# pip install openpyxl python-docx

from openpyxl import load_workbook
from docx import Document
from datetime import datetime

planilha_fornecedores = load_workbook('./fornecedores.xlsx')