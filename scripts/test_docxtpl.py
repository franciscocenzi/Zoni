import sys
from docxtpl import DocxTemplate

try:
    doc = DocxTemplate("c:/Users/franciscocenzi/Documents/GitHub/Zoni/infraestrutura/relatorios/modelos/modelo_relatorio.docx")
    context = {"DADOS_CADASTRAIS_LIST": []}
    doc.render(context)
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {e}")
