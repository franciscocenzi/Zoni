import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

def limpar_documento(doc):
    """Remove todo o corpo de texto (parágrafos e tabelas), mantendo Headers, Footers e Estilos Base."""
    # Deleta parágrafos de trás pra frente
    for paragraph in reversed(doc.paragraphs):
        p_element = paragraph._element
        p_element.getparent().remove(p_element)
        paragraph._p = paragraph._element = None
    
    # Deleta tabelas
    for table in reversed(doc.tables):
        t_element = table._element
        t_element.getparent().remove(t_element)
        table._tbl = table._element = None

def aplicar_estilo_amigavel(doc):
    """Garante que exista estilo Heading 1 e Normal limpos para o Zôni."""
    styles = doc.styles
    try:
        if 'ZoniHeading' not in styles:
            style = styles.add_style('ZoniHeading', WD_STYLE_TYPE.PARAGRAPH)
            style.base_style = styles['Heading 1'] if 'Heading 1' in styles else styles['Normal']
            font = style.font
            font.name = 'Arial'
            font.size = Pt(12)
            font.bold = True
            font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass

def adicionar_tabela_jinja(doc, cols, row_headers, list_var_name, fields_mapped):
    """Cria uma tabela no Word com um loop jinja mágico."""
    table = doc.add_table(rows=2, cols=cols)
    table.style = 'Table Grid'
    
    # Cabeçalho
    hdr_cells = table.rows[0].cells
    for i, hc in enumerate(row_headers):
        hdr_cells[i].text = hc
        # Negrito no cabeçalho
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                
    # Loop Jinja na primeira célula da linha 1
    row_cells = table.rows[1].cells
    row_cells[0].text = f"{{% tr for row in {list_var_name} %}}{{{{{ fields_mapped[0] }}}}}"
    for i in range(1, cols):
        if i == cols - 1:
            row_cells[i].text = f"{{{{{ fields_mapped[i] }}}}}{{% tr endfor %}}"
        else:
            row_cells[i].text = f"{{{{{ fields_mapped[i] }}}}}"
    return table

def construir_modelo_zoni(doc):
    aplicar_estilo_amigavel(doc)
    style_h = 'ZoniHeading' if 'ZoniHeading' in doc.styles else 'Heading 1'
    
    # Título Principal
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run("RELATÓRIO TÉCNICO DE ANÁLISE URBANÍSTICA")
    r.bold = True
    r.font.size = Pt(14)
    r.underline = True
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p_sub.add_run("{{ TIPO_ANALISE }}")
    r2.bold = True
    r2.font.size = Pt(12)
    
    doc.add_paragraph("Emissão: {{ DATA_COMPLETA }} às {{ HORA }}", style='Normal')
    doc.add_paragraph()
    
    # 1. DADOS CADASTRAIS
    doc.add_paragraph("1. DADOS CADASTRAIS", style=style_h)
    adicionar_tabela_jinja(doc, 5, 
        ["Proprietário", "Inscrição", "Endereço Completo", "Loteamento", "Área (m²)"],
        "DADOS_CADASTRAIS_LIST",
        ["row.proprietario", "row.inscricao", "row.endereco", "row.loteamento", "row.area"]
    )
    doc.add_paragraph()
    
    # 2. LIMITES DO TERRENO
    doc.add_paragraph("2. LIMITES DO TERRENO (TESTADAS E DIVISAS)", style=style_h)
    adicionar_tabela_jinja(doc, 2, 
        ["Limite / Logradouro", "Comprimento (m)"],
        "TABELA_TESTADAS_LIST",
        ["row.limite", "row.comprimento"]
    )
    doc.add_paragraph()
    
    # 3. ZONEAMENTO E ÍNDICES
    doc.add_paragraph("3. ZONEAMENTO INCIDENTE", style=style_h)
    # Tabela estática de índices
    t_z = doc.add_table(rows=6, cols=2)
    t_z.style = 'Table Grid'
    pares = [
        ("C.A. Máximo", "{{ CA_MAX_AJ }}"),
        ("C.A. Básico", "{{ CA_BAS }}"),
        ("Taxa Permeabilidade (TPS)", "{{ TPS }}"),
        ("Taxa de Ocupação (TOS)", "{{ TOS }}"),
        ("Recuo Frontal", "{{ RF }}"),
        ("Nº de Pavimentos", "{{ NP_BAS }}"),
    ]
    for i, (k, v) in enumerate(pares):
        t_z.rows[i].cells[0].text = k
        t_z.rows[i].cells[1].text = v
    doc.add_paragraph()
    
    # 4. ÁREAS DE PRESERVAÇÃO PERMANENTE (APP)
    doc.add_paragraph("4. ÁREAS DE PRESERVAÇÃO PERMANENTE (APP)", style=style_h)
    doc.add_paragraph("Faixa NUIC: {{ APP_FAIXA_STATUS }} ({{ APP_FAIXA_LARGURA }}) - {{ APP_FAIXA_OBS }}")
    doc.add_paragraph("Manguezal: {{ APP_MANGUE_STATUS }} - {{ APP_MANGUE_OBS }}")
    doc.add_paragraph()
    
    # 5. RISCOS GEOAMBIENTAIS
    doc.add_paragraph("5. RISCOS GEOAMBIENTAIS", style=style_h)
    doc.add_paragraph("Risco de Inundação: {{ RISCO_INUND_CLASSE }} - Grau: {{ RISCO_INUND_GRAU }}")
    doc.add_paragraph("Recomendação: {{ RISCO_INUND_RECOM }}")
    doc.add_paragraph("Risco Movimento Massa: {{ RISCO_MOV_CLASSE }} - Grau: {{ RISCO_MOV_GRAU }}")
    doc.add_paragraph("Recomendação: {{ RISCO_MOV_RECOM }}")
    doc.add_paragraph()
    
    # 6. INCLINAÇÃO
    doc.add_paragraph("6. INCLINAÇÃO DO TERRENO", style=style_h)
    adicionar_tabela_jinja(doc, 4, 
        ["Faixa de Inclinação", "Área (m²)", "% da Área", "Notas"],
        "TABELA_INCLINACAO_LIST",
        ["row.faixa", "row.area", "row.perc", "row.notas"]
    )
    doc.add_paragraph()

    # 7. MAPAS / IMAGENS (Anexos futuros mantendo a formatação padrão)
    p_img = doc.add_paragraph("7. MAPA DE SITUAÇÃO / ANEXOS GRAFICOS", style=style_h)
    p_img_placeholder = doc.add_paragraph()
    p_img_placeholder.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Tag auxiliar inline pro docxtpl (podemos injetar objeto InlineImage se tiver {{ IMAGEM_MAPA }})
    p_img_placeholder.add_run("{{ IMAGEM_MAPA }}").italic = True
    
def main():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_file = os.path.join(base_dir, "2024.005135_1 - RAPS - Loteamento Cidade dos Lagos.docx")
    out_dir = os.path.join(base_dir, "infraestrutura", "relatorios", "modelos")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "modelo_relatorio.docx")
    
    print(f"Lendo base corporativa: {input_file}")
    doc = Document(input_file)
    limpar_documento(doc)
    construir_modelo_zoni(doc)
    
    doc.save(out_file)
    print(f"Modelo salvo em: {out_file}")

if __name__ == "__main__":
    main()
