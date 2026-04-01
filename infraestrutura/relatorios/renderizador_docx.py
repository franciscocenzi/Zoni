"""
Renderizador nativo DOCX para Zôni v2.
Gera o relatório programaticamente via python-docx, preservando o cabeçalho/rodapé corporativo.
Paleta: tons de cinza. Tabelas ajustadas para A4 útil (~15,5cm).
"""
import os
import sys
import subprocess
import re
from datetime import datetime
from typing import List

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
# Paleta (cinzas)
# ──────────────────────────────────────────────────────────────────
COR_HDR_BG   = "595959"   # cinza escuro — fundo do cabeçalho de tabela
COR_HDR_FG   = RGBColor(0xFF, 0xFF, 0xFF)  # branco — texto do cabeçalho
COR_ZEBRA_A  = "F3F3F3"   # cinza clarinho — linha par
COR_ZEBRA_B  = "FFFFFF"   # branco — linha ímpar
COR_TITULO   = RGBColor(0x40, 0x40, 0x40)  # cinza escuro — títulos de seção
COR_SEP      = "CCCCCC"   # cinza médio — linhas divisórias (não usada por ora)

# Cores semafóricas de risco (mantidas leves)
COR_RISCO = {
    "ALTA":       ("FFCCCC", "Exige Estudo Hidrológico/Hidráulico ou Geotécnico completo."),
    "ALTO":       ("FFCCCC", "Exige Estudo Hidrológico/Hidráulico ou Geotécnico completo."),
    "MÉDIA":      ("FFF3CD", "Recomenda-se investigação preliminar."),
    "MEDIA":      ("FFF3CD", "Recomenda-se investigação preliminar."),
    "BAIXA":      ("D4EDDA", "Procedimentos construtivos padrão geralmente suficientes."),
    "BAIXO":      ("D4EDDA", "Procedimentos construtivos padrão geralmente suficientes."),
    "MUITO BAIXA":("D4EDDA", "Muito baixa suscetibilidade."),
}

# Largura útil A4 (margens corporativas 3cm+2.5cm ≈ 15.5cm)
LARGURA_UTIL = 15.5


# ──────────────────────────────────────────────────────────────────
# Helpers de formatação
# ──────────────────────────────────────────────────────────────────
def _s(v) -> str:
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
    """Define cor de fundo de célula via XML."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.replace("#", "").upper())
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    tcPr.append(shd)

def _dados_risco(classe: str):
    s = str(classe).upper()
    for k, (cor, recom) in COR_RISCO.items():
        if k in s:
            return cor, classe if classe != "-" else "Não classificado", recom
    return "EEEEEE", "Não classificado" if classe in ("-", "") else classe, "-"


# ──────────────────────────────────────────────────────────────────
# Construtores de bloco
# ──────────────────────────────────────────────────────────────────
def _add_heading(doc, texto: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(texto.upper())
    run.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = COR_TITULO
    return p


def _add_table(doc, headers: List[str], rows_data: List[List[str]], widths_cm: List[float]):
    """
    Monta uma tabela simples com cabeçalho cinza escuro, linhas zebradas e
    larguras fixas. A soma de widths_cm deve ser <= LARGURA_UTIL.
    """
    n_cols = len(headers)
    n_rows = max(1 + len(rows_data), 2)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = 'Table Grid'
    table.autofit = False

    # Cabeçalho
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        set_cell_color(cell, COR_HDR_BG)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(8)
                r.font.color.rgb = COR_HDR_FG

    # Dados
    if rows_data:
        for ri, row_vals in enumerate(rows_data):
            row = table.rows[ri + 1]
            bg = COR_ZEBRA_A if ri % 2 == 0 else COR_ZEBRA_B
            for ci, val in enumerate(row_vals[:n_cols]):
                row.cells[ci].text = _s(val)
                set_cell_color(row.cells[ci], bg)
                for p in row.cells[ci].paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8)
    else:
        row = table.rows[1]
        row.cells[0].text = "Sem dados."

    # Larguras fixas
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            if ci < len(widths_cm):
                cell.width = Cm(widths_cm[ci])

    return table


# ──────────────────────────────────────────────────────────────────
# Seções do relatório
# ──────────────────────────────────────────────────────────────────
def _sec_dados_cadastrais(doc, ctx):
    _add_heading(doc, "1. Dados Cadastrais")
    ident = ctx.get("identificacao") or {}
    ident_list = ident if isinstance(ident, list) else [ident]
    rows = []
    for d in ident_list:
        insc = _s(d.get("inscricao_imobiliaria"))
        cad  = _s(d.get("numero_cadastral"))
        ids  = " / ".join(x for x in [insc, cad] if x != "-") or "-"
        logr = _s(d.get("logradouro"))
        num  = _s(d.get("numero", "S/N"))
        bairro = _s(d.get("bairro"))
        end_parts = []
        if logr != "-": end_parts.append(f"{logr}, {num}")
        if bairro != "-": end_parts.append(f"Bairro {bairro}")
        endereco = " — ".join(end_parts) or "-"
        lot = _s(d.get("loteamento"))
        qd  = _s(d.get("quadra"))
        lt  = _s(d.get("lote"))
        lot_str = " | ".join(x for x in [
            lot if lot != "-" else None,
            f"Qd {qd}" if qd != "-" else None,
            f"Lt {lt}" if lt != "-" else None
        ] if x) or "-"
        area = _fmt_float(d.get("area_m2")) if d.get("area_m2") else "-"
        rows.append([_s(d.get("proprietario")), ids, endereco, lot_str, area])

    # 3.5 + 2.5 + 4.5 + 3.0 + 2.0 = 15.5
    _add_table(doc,
        ["Proprietário", "Inscrição/Cad.", "Endereço", "Loteamento/Qd/Lt", "Área (m²)"],
        rows, widths_cm=[3.5, 2.5, 4.5, 3.0, 2.0]
    )
    doc.add_paragraph()


def _sec_testadas(doc, ctx):
    _add_heading(doc, "2. Limites do Terreno (Testadas e Divisas)")
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
            log  = _s(s.get("logradouro"))
            conf = _s(s.get("confrontante"))
            comp = _fmt_float(s.get("comprimento_m"))
            desc = f"TESTADA — {log}" if tipo == "TESTADA" else f"DIVISA — {conf}"
            rows.append([desc, comp])

    # 11.5 + 4.0 = 15.5
    _add_table(doc,
        ["Limite / Logradouro / Confrontante", "Comprimento (m)"],
        rows, widths_cm=[11.5, 4.0]
    )
    doc.add_paragraph()


def _sec_zoneamento(doc, ctx):
    _add_heading(doc, "3. Zoneamento Incidente")
    zr     = ctx.get("zoneamento_resolvido") or {}
    zonas  = zr.get("zonas") or []
    rows   = []
    if zonas:
        for z in zonas:
            param  = z.get("parametros") or {}
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
        indices = ctx.get("indices") or {}
        param   = indices.get("parametros") or {}
        extras  = param.get("extras") or {}
        rows.append([
            z_nome, _fmt_float(ctx.get("area_lote_m2")), "100",
            _fmt_float(param.get("CA_max")), _fmt_float(param.get("CA_bas")),
            _fmt_perc(param.get("Tperm")), _fmt_perc(param.get("Tocup")),
            _s(param.get("Npav_bas")), _s(param.get("Npav_max")),
            _s(extras.get("RF") or extras.get("RF_Sec")),
        ])

    # 2.0+2.0+1.1+1.2+1.2+1.2+1.2+1.3+1.3+2.0 = 15.5
    _add_table(doc,
        ["Zona", "Área (m²)", "%", "CA máx", "CA bas", "TPS", "TOS", "Pav Bas", "Pav Máx", "Recuo Fr"],
        rows,
        widths_cm=[2.0, 2.0, 1.1, 1.2, 1.2, 1.2, 1.2, 1.3, 1.3, 2.0]
    )
    doc.add_paragraph()


def _sec_app(doc, ctx):
    _add_heading(doc, "4. Áreas de Preservação Permanente (APP)")
    amb = ctx.get("ambiente") or {}
    em_nuic   = amb.get("em_app_faixa_nuic")
    largura   = _s(amb.get("largura_faixa_m"))
    notas     = amb.get("notas") or []
    obs_nuic  = "; ".join(str(n) for n in notas[:2]) if em_nuic and notas else (
                "Sem curso d'água identificado no lote." if not em_nuic else "")
    em_mangue = amb.get("em_app_manguezal")
    obs_mangue= "; ".join(str(n) for n in notas[2:4]) if em_mangue and len(notas) > 2 else (
                "Sem manguezal identificado." if not em_mangue else "Manguezal detectado.")

    larg_str = f"{largura} m" if em_nuic and largura != "-" else "-"

    # 3.5 + 3.0 + 9.0 = 15.5
    _add_table(doc,
        ["Tipo de APP", "Situação / Largura", "Observações"],
        [
            ["Faixa Marginal (NUIC)", f"{'Presente' if em_nuic else 'Ausente'} — {larg_str}", obs_nuic],
            ["Manguezal",             "Presente" if em_mangue else "Ausente",                  obs_mangue],
        ],
        widths_cm=[3.5, 3.0, 9.0]
    )
    doc.add_paragraph()


def _sec_risco(doc, ctx):
    _add_heading(doc, "5. Riscos Geoambientais")
    risco = ctx.get("risco") or {}
    cl_i  = _s(risco.get("classe_inundacao"))
    cl_m  = _s(risco.get("classe_movimento_massa"))

    cor_i, grau_i, recom_i = _dados_risco(cl_i)
    cor_m, grau_m, recom_m = _dados_risco(cl_m)

    table = doc.add_table(rows=3, cols=4)
    table.style = 'Table Grid'
    table.autofit = False

    # Linha 0: cabeçalhos cinza
    table.rows[0].cells[0].merge(table.rows[0].cells[1])
    table.rows[0].cells[0].text = "Suscetibilidade a Inundação"
    table.rows[0].cells[2].merge(table.rows[0].cells[3])
    table.rows[0].cells[2].text = "Suscetibilidade a Movimentos de Massa"
    for cell in [table.rows[0].cells[0], table.rows[0].cells[2]]:
        set_cell_color(cell, COR_HDR_BG)
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold = True; r.font.size = Pt(9)
                r.font.color.rgb = COR_HDR_FG

    # Linha 1: quadradinho de cor + grau
    table.rows[1].cells[0].text = "   "
    set_cell_color(table.rows[1].cells[0], cor_i)
    table.rows[1].cells[1].text = grau_i
    for r in table.rows[1].cells[1].paragraphs[0].runs:
        r.font.bold = True; r.font.size = Pt(10)
    set_cell_color(table.rows[1].cells[1], COR_ZEBRA_A)

    table.rows[1].cells[2].text = "   "
    set_cell_color(table.rows[1].cells[2], cor_m)
    table.rows[1].cells[3].text = grau_m
    for r in table.rows[1].cells[3].paragraphs[0].runs:
        r.font.bold = True; r.font.size = Pt(10)
    set_cell_color(table.rows[1].cells[3], COR_ZEBRA_A)

    # Linha 2: recomendações (fundo levemente cinza)
    table.rows[2].cells[0].merge(table.rows[2].cells[1])
    table.rows[2].cells[0].text = recom_i
    set_cell_color(table.rows[2].cells[0], COR_ZEBRA_B)
    table.rows[2].cells[2].merge(table.rows[2].cells[3])
    table.rows[2].cells[2].text = recom_m
    set_cell_color(table.rows[2].cells[2], COR_ZEBRA_B)
    for cell in [table.rows[2].cells[0], table.rows[2].cells[2]]:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(8)

    # Larguras: 1.0 + 6.75 + 1.0 + 6.75 = 15.5
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            cell.width = Cm([1.0, 6.75, 1.0, 6.75][ci] if ci < 4 else 1.0)

    doc.add_paragraph()


def _sec_inclinacao(doc, ctx):
    _add_heading(doc, "6. Inclinação do Terreno")
    incl  = ctx.get("inclinacao") or {}
    faixas = incl.get("faixas") or [] if isinstance(incl, dict) else []

    if not faixas:
        msg = incl.get("mensagem", "Análise de inclinação não disponível.") if isinstance(incl, dict) else "Não disponível."
        doc.add_paragraph(msg).runs[0].font.size = Pt(9)
        doc.add_paragraph()
        return

    table = doc.add_table(rows=1 + len(faixas), cols=5)
    table.style = 'Table Grid'
    table.autofit = False

    hdrs = ["Faixa de Inclinação", "Cor", "Área (m²)", "% da Área", "Notas"]
    for i, h in enumerate(hdrs):
        c = table.rows[0].cells[i]
        c.text = h
        set_cell_color(c, COR_HDR_BG)
        for p in c.paragraphs:
            for r in p.runs:
                r.font.bold = True; r.font.size = Pt(8)
                r.font.color.rgb = COR_HDR_FG

    for ri, f in enumerate(faixas):
        row  = table.rows[ri + 1]
        cor  = str(f.get("cor", "#CCCCCC")).replace("#", "").upper()
        app  = "APP" if f.get("app") else "-"
        bg   = COR_ZEBRA_A if ri % 2 == 0 else COR_ZEBRA_B
        vals = [_s(f.get("faixa")), "", _fmt_float(f.get("area_m2")),
                _fmt_float(f.get("percentual"), dec=1) + "%", app]
        for ci, val in enumerate(vals):
            row.cells[ci].text = val
            set_cell_color(row.cells[ci], bg)
            for p in row.cells[ci].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8)
        # célula de cor
        set_cell_color(row.cells[1], cor)

    # 4.5 + 1.5 + 3.5 + 3.5 + 2.5 = 15.5 (mas 2 cols ficam com index correto)
    widths = [4.5, 1.5, 3.5, 3.5, 2.5]
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            if ci < len(widths):
                cell.width = Cm(widths[ci])

    # Totais APP se houver
    area_app = incl.get("area_app_inclinacao_m2") if isinstance(incl, dict) else None
    if area_app and float(area_app) > 0:
        p = doc.add_paragraph(f"• Total em APP por inclinação (>45°): {_fmt_float(area_app)} m²")
        p.runs[0].font.size = Pt(8)

    doc.add_paragraph()


def _sec_notas(doc, ctx):
    _add_heading(doc, "7. Notas e Condicionantes Técnicas")

    try:
        from infraestrutura.relatorios.renderizador_html import _montar_listas_notas_separadas
        listas = _montar_listas_notas_separadas(ctx)
    except Exception:
        listas = {}

    def _add_lista(titulo, chave, fallback):
        html = listas.get(chave, "")
        itens = []
        if html and isinstance(html, str):
            itens = [re.sub('<[^>]+>', '', i).strip()
                     for i in re.findall(r'<li[^>]*>(.*?)</li>', html, re.S)
                     if i.strip()]
        p = doc.add_paragraph()
        p.add_run(titulo).bold = True
        if itens:
            for item in itens:
                b = doc.add_paragraph(style='List Bullet')
                b.add_run(item).font.size = Pt(8)
        else:
            q = doc.add_paragraph(fallback)
            q.runs[0].font.size = Pt(9)

    _add_lista("Notas Técnicas e Legislativas:", "LISTA_NOTAS_ANEXO",
               "Nenhuma nota específica aplicada.")
    _add_lista("Condicionantes:", "LISTA_CONDICIONANTES",
               "Nenhuma condicionante identificada.")
    _add_lista("Restrições e Pendências:", "LISTA_RESTRICOES",
               "Nenhuma restrição crítica identificada.")
    doc.add_paragraph()


def _sec_mapa(doc):
    _add_heading(doc, "8. Mapa de Situação / Anexos Gráficos")
    p = doc.add_paragraph("[Mapa será inserido futuramente]")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)


# ──────────────────────────────────────────────────────────────────
# Classe principal
# ──────────────────────────────────────────────────────────────────
class RenderizadorDOCX:
    """Gera o relatório DOCX programaticamente usando python-docx."""

    def __init__(self):
        self.base_dir    = os.path.dirname(os.path.abspath(__file__))
        self.modelo_path = os.path.join(self.base_dir, "modelos", "modelo_relatorio.docx")

    def renderizar_e_salvar(self, contexto: dict, caminho_saida: str) -> tuple:
        try:
            if os.path.exists(self.modelo_path):
                doc = Document(self.modelo_path)
                for p in list(doc.paragraphs):
                    p._element.getparent().remove(p._element)
                for t in list(doc.tables):
                    t._element.getparent().remove(t._element)
            else:
                doc = Document()

            agora = datetime.now()
            tipo  = "Gleba Unificada" if contexto.get("area_gleba_unificada") else "Lote"

            # Título
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("RELATÓRIO TÉCNICO DE ANÁLISE URBANÍSTICA")
            r.bold = True; r.font.size = Pt(12); r.underline = True
            r.font.color.rgb = COR_TITULO

            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(tipo)
            r2.bold = True; r2.font.size = Pt(10); r2.font.color.rgb = COR_TITULO

            pe = doc.add_paragraph(
                f"Emissão: {agora.strftime('%d/%m/%Y')} às {agora.strftime('%H:%M')}")
            pe.runs[0].font.size = Pt(9); pe.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            doc.add_paragraph()

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
