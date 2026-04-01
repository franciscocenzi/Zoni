"""
Renderizador nativo DOCX para Zôni v2.
Gera o relatório programaticamente via python-docx, preservando o cabeçalho/rodapé corporativo.
Sem Jinja2/docxtpl — evita todos os problemas de parsing de templates.
"""
import os
import sys
import subprocess
from datetime import datetime
from typing import Dict, Any, List

from qgis.core import QgsMessageLog, Qgis

# ──────────────────────────────────────────────────────────────────
# Dependências
# ──────────────────────────────────────────────────────────────────
try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    QgsMessageLog.logMessage("python-docx ausente. Instalando...", "Zôni v2", Qgis.Warning)
    try:
        python_exe = os.path.join(sys.prefix, 'python.exe')
        if not os.path.exists(python_exe):
            python_exe = "python"
        subprocess.check_call([python_exe, "-m", "pip", "install", "python-docx"])
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except Exception as e:
        QgsMessageLog.logMessage(f"Falha ao instalar python-docx: {e}", "Zôni v2", Qgis.Critical)


# ──────────────────────────────────────────────────────────────────
# Helpers de formatação
# ──────────────────────────────────────────────────────────────────
AZUL_ZONI = RGBColor(0x1F, 0x49, 0x7D)

def _s(v) -> str:
    """Converte None para '-', qualquer outro valor para string."""
    if v is None or str(v).strip() in ("", "None", "none"):
        return "-"
    return str(v).strip()

def _fmt_float(v, dec=2) -> str:
    try:
        f = float(str(v).replace(",", "."))
        if f == int(f):
            return f"{int(f):,}".replace(",", ".")
        return f"{f:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return _s(v)

def _fmt_perc(v) -> str:
    try:
        f = float(str(v).replace(",", ".").replace("%", ""))
        if 0 < f <= 1.0:
            f *= 100
        return f"{f:.1f}%".replace(".", ",")
    except Exception:
        return _s(v)

def set_cell_color(cell, hex_color: str):
    """Define a cor de fundo de uma célula (hex sem #)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.replace("#", "").upper())
    # Remove shd anterior se existir
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    tcPr.append(shd)

def _add_heading(doc, texto: str, nivel=1):
    """Adiciona um parágrafo de seção estilizado."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = AZUL_ZONI
    return p

def _add_table(doc, headers: List[str], rows_data: List[List[str]], widths_cm: List[float] = None):
    """
    Adiciona uma tabela com cabeçalho azul e linhas de dados.
    rows_data: lista de listas de strings.
    """
    n_cols = len(headers)
    n_rows = 1 + len(rows_data) if rows_data else 2
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = 'Table Grid'
    table.autofit = False

    # Cabeçalho
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        set_cell_color(cell, "1F497D")  # azul corporativo
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Dados
    if rows_data:
        for ri, row_vals in enumerate(rows_data):
            row = table.rows[ri + 1]
            bg = "F2F7FF" if ri % 2 == 0 else "FFFFFF"
            for ci, val in enumerate(row_vals):
                if ci < n_cols:
                    row.cells[ci].text = _s(val)
                    set_cell_color(row.cells[ci], bg)
                    for p in row.cells[ci].paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(8)
    else:
        row = table.rows[1]
        row.cells[0].text = "Sem dados."

    # Larguras
    if widths_cm:
        for row in table.rows:
            for ci, cell in enumerate(row.cells):
                if ci < len(widths_cm):
                    cell.width = Cm(widths_cm[ci])

    return table


# ──────────────────────────────────────────────────────────────────
# Construtores de seção
# ──────────────────────────────────────────────────────────────────

def _sec_dados_cadastrais(doc, ctx):
    _add_heading(doc, "1. DADOS CADASTRAIS")
    ident = ctx.get("identificacao") or {}
    ident_list = ident if isinstance(ident, list) else [ident]
    rows = []
    for d in ident_list:
        insc = _s(d.get("inscricao_imobiliaria"))
        cad = _s(d.get("numero_cadastral"))
        ids = " / ".join(x for x in [insc, cad] if x != "-") or "-"
        logr = _s(d.get("logradouro"))
        num = _s(d.get("numero", "S/N"))
        bairro = _s(d.get("bairro"))
        end_parts = []
        if logr != "-": end_parts.append(f"{logr}, {num}")
        if bairro != "-": end_parts.append(f"Bairro {bairro}")
        endereco = " — ".join(end_parts) or "-"
        lot = _s(d.get("loteamento"))
        qd = _s(d.get("quadra"))
        lt = _s(d.get("lote"))
        lot_str = " | ".join(x for x in [
            lot if lot != "-" else None,
            f"Qd: {qd}" if qd != "-" else None,
            f"Lt: {lt}" if lt != "-" else None
        ] if x) or "-"
        area = _fmt_float(d.get("area_m2")) if d.get("area_m2") else "-"
        rows.append([_s(d.get("proprietario")), ids, endereco, lot_str, area])

    _add_table(doc,
        ["Proprietário", "Inscrição / Cad.", "Endereço", "Loteamento / Qd / Lt", "Área (m²)"],
        rows, widths_cm=[4.5, 3.0, 4.5, 3.5, 2.5]
    )
    doc.add_paragraph()


def _sec_testadas(doc, ctx):
    _add_heading(doc, "2. LIMITES DO TERRENO (TESTADAS E DIVISAS)")
    segs = ctx.get("segmentos_limites") or []
    testadas_log = ctx.get("testadas_por_logradouro") or {}
    confrontantes = ctx.get("confrontantes_por_proprietario") or {}

    rows = []
    if isinstance(testadas_log, dict) and testadas_log:
        for log, comp in testadas_log.items():
            rows.append([f"TESTADA — {log}", _fmt_float(comp)])
    if isinstance(confrontantes, dict) and confrontantes:
        for prop, comp in confrontantes.items():
            rows.append([f"DIVISA — {prop}", _fmt_float(comp)])
    if not rows and isinstance(segs, list):
        for s in segs:
            tipo = (s.get("tipo_limite") or "").upper()
            log = _s(s.get("logradouro"))
            conf = _s(s.get("confrontante"))
            comp = _fmt_float(s.get("comprimento_m"))
            if tipo == "TESTADA":
                desc = f"TESTADA — {log}"
            else:
                desc = f"DIVISA — {conf}"
            rows.append([desc, comp])

    _add_table(doc,
        ["Limite / Logradouro / Confrontante", "Comprimento (m)"],
        rows, widths_cm=[13.5, 3.5]
    )
    doc.add_paragraph()


def _sec_zoneamento(doc, ctx):
    _add_heading(doc, "3. ZONEAMENTO INCIDENTE")
    zr = ctx.get("zoneamento_resolvido") or {}
    zonas = zr.get("zonas") or []
    rows = []
    if zonas:
        for z in zonas:
            param = z.get("parametros") or {}
            extras = param.get("extras") or {}
            rows.append([
                _s(z.get("codigo")),
                _fmt_float(z.get("area_m2")),
                _fmt_float(z.get("percentual_area"), dec=1),
                _fmt_float(param.get("CA_max")),
                _fmt_float(param.get("CA_bas")),
                _fmt_perc(param.get("Tperm")),
                _fmt_perc(param.get("Tocup")),
                _s(param.get("Npav_bas")),
                _s(param.get("Npav_max")),
                _s(extras.get("RF") or extras.get("RF_Sec")),
            ])
    else:
        z_nome = (ctx.get("zoneamento") or {}).get("zona", "-")
        indices = (ctx.get("indices") or {})
        param = indices.get("parametros") or {}
        extras = param.get("extras") or {}
        rows.append([
            z_nome, _fmt_float(ctx.get("area_lote_m2")), "100",
            _fmt_float(param.get("CA_max")), _fmt_float(param.get("CA_bas")),
            _fmt_perc(param.get("Tperm")), _fmt_perc(param.get("Tocup")),
            _s(param.get("Npav_bas")), _s(param.get("Npav_max")),
            _s(extras.get("RF") or extras.get("RF_Sec")),
        ])

    _add_table(doc,
        ["Zona", "Área (m²)", "%", "CA máx", "CA bas", "TPS", "TOS", "Pav Bas", "Pav Máx", "Recuo Fr"],
        rows, widths_cm=[2.2, 2.2, 1.4, 1.4, 1.4, 1.4, 1.4, 1.5, 1.5, 2.0]
    )
    doc.add_paragraph()


def _sec_app(doc, ctx):
    _add_heading(doc, "4. ÁREAS DE PRESERVAÇÃO PERMANENTE (APP)")
    amb = ctx.get("ambiente") or {}
    em_nuic = amb.get("em_app_faixa_nuic")
    largura = _s(amb.get("largura_faixa_m"))
    notas = amb.get("notas") or []
    obs_nuic = "; ".join(notas[:2]) if em_nuic and notas else ("Sem curso d'água identificado." if not em_nuic else "")
    em_mangue = amb.get("em_app_manguezal")
    obs_mangue = "; ".join(notas[2:4]) if em_mangue and len(notas) > 2 else ("Sem manguezal identificado." if not em_mangue else "Manguezal detectado.")

    status_nuic = "Presente" if em_nuic else "Ausente"
    status_mangue = "Presente" if em_mangue else "Ausente"

    _add_table(doc,
        ["Tipo de APP", "Situação", "Informações"],
        [
            ["Faixa Marginal (NUIC)", f"{status_nuic} — {largura} m", obs_nuic],
            ["Manguezal", status_mangue, obs_mangue],
        ],
        widths_cm=[4.0, 3.5, 10.5]
    )
    doc.add_paragraph()


def _sec_risco(doc, ctx):
    _add_heading(doc, "5. RISCOS GEOAMBIENTAIS")
    risco = ctx.get("risco") or {}
    classe_inund = _s(risco.get("classe_inundacao"))
    classe_mov = _s(risco.get("classe_movimento_massa"))

    COR_MAP = {
        "ALTA":       ("FFCCCC", "Área com alta suscetibilidade. Exige EHH detalhado ou laudo geotécnico."),
        "ALTA":       ("FFCCCC", "Área com alta suscetibilidade. Exige EHH detalhado ou laudo geotécnico."),
        "MÉDIA":      ("FFEB9C", "Suscetibilidade média. Recomenda-se investigação preliminar."),
        "MEDIA":      ("FFEB9C", "Suscetibilidade média. Recomenda-se investigação preliminar."),
        "BAIXA":      ("CCFFCC", "Baixa suscetibilidade. Procedimentos construtivos padrão."),
        "MUITO BAIXA":("CCFFCC", "Muito baixa suscetibilidade."),
    }

    def _dados_risco(classe):
        for k, (cor, recom) in COR_MAP.items():
            if k in classe.upper():
                return cor, classe, recom
        return "F2F2F2", classe if classe != "-" else "Não classificado", "-"

    cor_i, grau_i, recom_i = _dados_risco(classe_inund)
    cor_m, grau_m, recom_m = _dados_risco(classe_mov)

    table = doc.add_table(rows=3, cols=4)
    table.style = 'Table Grid'
    table.autofit = False

    # Linha 0: cabeçalhos
    table.rows[0].cells[0].merge(table.rows[0].cells[1])
    table.rows[0].cells[0].text = "Suscetibilidade a Inundação"
    table.rows[0].cells[2].merge(table.rows[0].cells[3])
    table.rows[0].cells[2].text = "Suscetibilidade a Movimentos de Massa"
    for cell in [table.rows[0].cells[0], table.rows[0].cells[2]]:
        set_cell_color(cell, "1F497D")
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold = True; r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)

    # Linha 1: cor + grau
    table.rows[1].cells[0].text = ""  # célula colorida
    set_cell_color(table.rows[1].cells[0], cor_i)
    table.rows[1].cells[1].text = grau_i
    for r in table.rows[1].cells[1].paragraphs[0].runs:
        r.font.bold = True; r.font.size = Pt(10)

    table.rows[1].cells[2].text = ""  # célula colorida
    set_cell_color(table.rows[1].cells[2], cor_m)
    table.rows[1].cells[3].text = grau_m
    for r in table.rows[1].cells[3].paragraphs[0].runs:
        r.font.bold = True; r.font.size = Pt(10)

    # Linha 2: recomendações
    table.rows[2].cells[0].merge(table.rows[2].cells[1])
    table.rows[2].cells[0].text = recom_i
    table.rows[2].cells[2].merge(table.rows[2].cells[3])
    table.rows[2].cells[2].text = recom_m
    set_cell_color(table.rows[2].cells[0], "FFF9E6")
    set_cell_color(table.rows[2].cells[2], "FFF9E6")

    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            w = Cm(1.0) if ci in (0, 2) else Cm(7.5)
            cell.width = w

    doc.add_paragraph()


def _sec_inclinacao(doc, ctx):
    _add_heading(doc, "6. INCLINAÇÃO DO TERRENO")
    incl = ctx.get("inclinacao") or {}
    faixas = incl.get("faixas") or [] if isinstance(incl, dict) else []

    if not faixas:
        doc.add_paragraph(incl.get("mensagem", "Análise de inclinação não disponível.") if isinstance(incl, dict) else "Análise de inclinação não disponível.")
        doc.add_paragraph()
        return

    n_cols = 5
    table = doc.add_table(rows=1 + len(faixas), cols=n_cols)
    table.style = 'Table Grid'
    table.autofit = False

    hdrs = ["Faixa de Inclinação", "Cor", "Área (m²)", "% da Área", "Notas"]
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(hdrs):
        hdr_cells[i].text = h
        set_cell_color(hdr_cells[i], "1F497D")
        for p in hdr_cells[i].paragraphs:
            for r in p.runs:
                r.font.bold = True; r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(255, 255, 255)

    for ri, f in enumerate(faixas):
        row = table.rows[ri + 1]
        cor_hex = str(f.get("cor", "#CCCCCC")).replace("#", "").upper()
        app_flag = "APP" if f.get("app") else "-"
        row.cells[0].text = _s(f.get("faixa"))
        row.cells[1].text = ""  # célula colorida
        set_cell_color(row.cells[1], cor_hex)
        row.cells[2].text = _fmt_float(f.get("area_m2"))
        row.cells[3].text = _fmt_float(f.get("percentual"), dec=1) + "%"
        row.cells[4].text = app_flag
        bg = "F2F7FF" if ri % 2 == 0 else "FFFFFF"
        for ci in [0, 2, 3, 4]:
            set_cell_color(row.cells[ci], bg)
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8)

    widths = [4.5, 1.2, 3.0, 2.5, 2.0]
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            if ci < len(widths):
                cell.width = Cm(widths[ci])

    doc.add_paragraph()


def _sec_notas(doc, ctx):
    _add_heading(doc, "7. NOTAS E CONDICIONANTES TÉCNICAS")
    # Reutiliza a lógica do HTML original para montar as listas
    try:
        from infraestrutura.relatorios.renderizador_html import (
            _montar_listas_notas_separadas
        )
        listas = _montar_listas_notas_separadas(ctx)
    except Exception:
        listas = {}

    def _add_lista(titulo, chave, fallback):
        html = listas.get(chave, "")
        if html and isinstance(html, str):
            import re
            itens = re.findall(r'<li[^>]*>(.*?)</li>', html, re.S)
            itens = [re.sub('<[^>]+>', '', i).strip() for i in itens if i.strip()]
        else:
            itens = []
        p = doc.add_paragraph()
        p.add_run(titulo).bold = True
        if itens:
            for item in itens:
                b = doc.add_paragraph(style='List Bullet')
                b.add_run(item).font.size = Pt(9)
        else:
            doc.add_paragraph(fallback).runs[0].font.size = Pt(9)

    _add_lista("Notas Técnicas e Legislativas:", "LISTA_NOTAS_ANEXO", "Nenhuma nota específica aplicada.")
    _add_lista("Condicionantes:", "LISTA_CONDICIONANTES", "Nenhuma condicionante identificada.")
    _add_lista("Restrições e Pendências:", "LISTA_RESTRICOES", "Nenhuma restrição crítica identificada.")
    doc.add_paragraph()


def _sec_mapa(doc):
    _add_heading(doc, "8. MAPA DE SITUAÇÃO / ANEXOS GRÁFICOS")
    p = doc.add_paragraph("[Mapa será inserido futuramente]")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True
    p.runs[0].font.size = Pt(9)


# ──────────────────────────────────────────────────────────────────
# Classe principal
# ──────────────────────────────────────────────────────────────────

class RenderizadorDOCX:
    """Gera o relatório DOCX programaticamente usando python-docx."""

    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.modelo_path = os.path.join(self.base_dir, "modelos", "modelo_relatorio.docx")

    def renderizar_e_salvar(self, contexto: dict, caminho_saida: str) -> tuple:
        try:
            # Abre o documento base para herdar cabeçalho/rodapé corporativo
            if os.path.exists(self.modelo_path):
                doc = Document(self.modelo_path)
                # Limpa o corpo mantendo header/footer
                for p in list(doc.paragraphs):
                    p._element.getparent().remove(p._element)
                for t in list(doc.tables):
                    t._element.getparent().remove(t._element)
            else:
                doc = Document()

            agora = datetime.now()
            tipo = "Gleba Unificada" if contexto.get("area_gleba_unificada") else "Lote"

            # Título
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("RELATÓRIO TÉCNICO DE ANÁLISE URBANÍSTICA")
            r.bold = True; r.font.size = Pt(13); r.font.color.rgb = AZUL_ZONI; r.underline = True

            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(tipo)
            r2.bold = True; r2.font.size = Pt(11)

            doc.add_paragraph(f"Emissão: {agora.strftime('%d/%m/%Y')} às {agora.strftime('%H:%M')}")
            doc.add_paragraph()

            # Seções
            _sec_dados_cadastrais(doc, contexto)
            _sec_testadas(doc, contexto)
            _sec_zoneamento(doc, contexto)
            _sec_app(doc, contexto)
            _sec_risco(doc, contexto)
            _sec_inclinacao(doc, contexto)
            _sec_notas(doc, contexto)
            _sec_mapa(doc)

            doc.save(caminho_saida)
            QgsMessageLog.logMessage(f"Relatório gerado: {caminho_saida}", "Zôni v2", Qgis.Success)
            return True, ""
        except Exception as e:
            QgsMessageLog.logMessage(f"Erro ao gerar DOCX: {e}", "Zôni v2", Qgis.Critical)
            import traceback
            QgsMessageLog.logMessage(traceback.format_exc(), "Zôni v2", Qgis.Critical)
            return False, str(e)
