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

def adicionar_tabela_jinja(doc, cols, row_headers, list_var_name, fields_mapped, style='Table Grid'):
    """Cria uma tabela no Word com um loop jinja mágico."""
    table = doc.add_table(rows=2, cols=cols)
    table.style = style
    
    # Cabeçalho
    hdr_cells = table.rows[0].cells
    for i, hc in enumerate(row_headers):
        hdr_cells[i].text = hc
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                
    def _render_field(field):
        """Envolve o campo em {{ }} a menos que já seja uma tag de bloco."""
        f = field.strip()
        if f.startswith('{%') or f.startswith('{{'):
            return field
        return '{{' + f + '}}'

    # Loop Jinja na primeira célula da linha 1
    row_cells = table.rows[1].cells
    row_cells[0].text = f"{{% for row in {list_var_name} %}}{_render_field(fields_mapped[0])}"
    for i in range(1, cols):
        if i == cols - 1:
            row_cells[i].text = f"{_render_field(fields_mapped[i])}{{% endfor %}}"
        else:
            row_cells[i].text = f"{_render_field(fields_mapped[i])}"
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
    adicionar_tabela_jinja(doc, 10, 
        ["Zona", "Área", "%", "C.A max", "C.A bas", "TPS (Perm)", "TOS (Ocup)", "Pav Bas", "Pav Max", "Recuo Fr"],
        "TABELA_ZONAS_LIST",
        ["row.codigo", "row.area", "row.perc", "row.ca_max", "row.ca_bas", "row.tps", "row.tos", "row.np_bas", "row.np_max", "row.rf"]
    )
    doc.add_paragraph()
    
    # 4. ÁREAS DE PRESERVAÇÃO PERMANENTE (APP)
    doc.add_paragraph("4. ÁREAS DE PRESERVAÇÃO PERMANENTE (APP)", style=style_h)
    doc.add_paragraph("Faixa NUIC: {{ APP_FAIXA_STATUS }} ({{ APP_FAIXA_LARGURA }} m) - {{ APP_FAIXA_OBS }}")
    doc.add_paragraph("Manguezal: {{ APP_MANGUE_STATUS }} - {{ APP_MANGUE_OBS }}")
    doc.add_paragraph()
    
    # 5. RISCOS GEOAMBIENTAIS
    doc.add_paragraph("5. RISCOS GEOAMBIENTAIS", style=style_h)
    
    t_r = doc.add_table(rows=3, cols=4)
    t_r.style = 'Normal Table'
    t_r.rows[0].cells[0].merge(t_r.rows[0].cells[1])
    t_r.rows[0].cells[0].text = "Suscetibilidade a Inundação"
    t_r.rows[0].cells[2].merge(t_r.rows[0].cells[3])
    t_r.rows[0].cells[2].text = "Suscetibilidade a Movimentos de Massa"
    for cell in t_r.rows[0].cells:
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                
    # Celula 1
    t_r.rows[1].cells[0].text = "{% cellbg RISCO_INUND_COR %} "
    t_r.rows[1].cells[1].text = "{{ RISCO_INUND_GRAU }}"
    t_r.rows[2].cells[0].merge(t_r.rows[2].cells[1])
    t_r.rows[2].cells[0].text = "{{ RISCO_INUND_RECOM }}"
    # Celula 2
    t_r.rows[1].cells[2].text = "{% cellbg RISCO_MOV_COR %} "
    t_r.rows[1].cells[3].text = "{{ RISCO_MOV_GRAU }}"
    t_r.rows[2].cells[2].merge(t_r.rows[2].cells[3])
    t_r.rows[2].cells[2].text = "{{ RISCO_MOV_RECOM }}"
    doc.add_paragraph()
    
    # 6. INCLINAÇÃO
    doc.add_paragraph("6. INCLINAÇÃO DO TERRENO", style=style_h)
    adicionar_tabela_jinja(doc, 5, 
        ["Faixa de Inclinação", "Legenda", "Área (m²)", "% da Área", "Notas"],
        "TABELA_INCLINACAO_LIST",
        ["row.faixa", "{% cellbg row.cor %} ", "row.area", "row.perc", "row.notas"],
        style="Normal Table"
    )
    doc.add_paragraph()
    
    # 7. NOTAS E CONDICIONANTES TÉCNICAS
    doc.add_paragraph("7. NOTAS E CONDICIONANTES TÉCNICAS", style=style_h)
    doc.add_paragraph("{{ LISTA_NOTAS_ANEXO_BULLETS }}")
    doc.add_paragraph("{{ LISTA_CONDICIONANTES_BULLETS }}")
    doc.add_paragraph("{{ LISTA_RESTRICOES_BULLETS }}")
    doc.add_paragraph()

    # 8. MAPAS / IMAGENS
    p_img = doc.add_paragraph("8. MAPA DE SITUAÇÃO / ANEXOS GRAFICOS", style=style_h)
    p_img_placeholder = doc.add_paragraph()
    p_img_placeholder.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Tag auxiliar inline pro docxtpl (podemos injetar objeto InlineImage se tiver {{ IMAGEM_MAPA }})
    p_img_placeholder.add_run("{{ IMAGEM_MAPA }}").italic = True
    
def main():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(base_dir, "infraestrutura", "relatorios", "modelos")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "modelo_relatorio.docx")
    
    # Usa o próprio documento atual como base de cabeçalhos
    input_file = out_file if os.path.exists(out_file) else None
    
    print(f"Lendo base corporativa: {input_file}")
    doc = Document(input_file)
    limpar_documento(doc)
    construir_modelo_zoni(doc)
    
    doc.save(out_file)
    print(f"Modelo salvo em: {out_file}")

if __name__ == "__main__":
    main()
