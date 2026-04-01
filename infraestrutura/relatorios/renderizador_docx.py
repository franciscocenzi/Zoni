"""
Renderizador nativo DOCX para Zôni v2 utilizando docxtpl.
Garanti a fidelidade visual utilizando um modelo Microsoft Word base.
"""
import os
import sys
import subprocess
from typing import Dict, Any
from qgis.core import QgsMessageLog, Qgis

# Auto-instalador de dependência nativa do Plugin
try:
    from docxtpl import DocxTemplate
except ImportError:
    QgsMessageLog.logMessage("Módulo docxtpl ausente. Iniciando instalação silenciosa no QGIS...", "Zôni v2", Qgis.Warning)
    try:
        # sys.executable retorna o qgis-bin.exe no Windows
        python_exe = os.path.join(sys.prefix, 'python.exe')
        if not os.path.exists(python_exe):
            python_exe = "python"
        subprocess.check_call([python_exe, "-m", "pip", "install", "docxtpl"])
        from docxtpl import DocxTemplate
        QgsMessageLog.logMessage("docxtpl instalado com sucesso.", "Zôni v2", Qgis.Success)
    except Exception as e:
        QgsMessageLog.logMessage(f"Falha ao instalar docxtpl via pip: {e}", "Zôni v2", Qgis.Critical)

def limpar_string(valor: Any) -> str:
    """Limpa a string removendo quebras vazias que possam quebrar o layout do Word."""
    if valor is None:
        return "-"
    return str(valor).strip()

def mapear_contexto_html_para_docx(contexto: Dict[str, Any]) -> Dict[str, Any]:
    """
    Traduz a saída focada em HTML antigo para variáveis planas aceitas pelo Jinja do docxtpl.
    Remove tags HTML brutas <tr>, <td>, <span> de listas e as transforma em dicionários estruturados.
    NOTA: Isso depende de injetar listas corretas de dicionários para iterar no Word.
    """
    ctx_docx = {}
    
    # Repassa variáveis escalares diretas
    for k, v in contexto.items():
        if isinstance(v, (str, int, float, bool)):
            ctx_docx[k] = limpar_string(v)

    # Identificações - Dicionário de Lote base
    # (Pega o fallback em string caso já esteja montado, mas o DocxTpl lida com formatadores puros)
    ctx_docx["N_LOTES"] = limpar_string(contexto.get("n_lotes"))
    ctx_docx["AREA_LOTE"] = limpar_string(contexto.get("area_lote_m2"))
    ctx_docx["N_TESTADAS"] = limpar_string(contexto.get("n_testadas"))
    ctx_docx["TESTADA_PRINCIPAL"] = limpar_string(contexto.get("testada_principal"))
    
    # Aqui precisamos receber os hashes processados originais se existirem 
    # ou usar os blocos em listas. Se o construtor já montou listas simples, passamos limpo:
    # EX: "LISTA_CONDICIONANTES" foi montado como strings com <li>. Precisamos limpar as tags <li>.
    
    for tag_key in ["LISTA_NOTAS_ANEXO", "LISTA_CONDICIONANTES", "LISTA_RESTRICOES", "LISTA_NOTAS"]:
        html_val = contexto.get(tag_key)
        if html_val and isinstance(html_val, str):
            linhas = html_val.replace("</li>", "").split("<li>")
            linhas_limpas = [linha.replace("<b>", "").replace("</b>", "").strip() for linha in linhas if linha.strip()]
        elif isinstance(html_val, list):
            linhas_limpas = [str(l) for l in html_val]
        else:
            linhas_limpas = []
        ctx_docx[tag_key + "_ARRAY"] = [{"texto": l} for l in linhas_limpas]
        # Versão pré-montada como string simples com bullet points (sem loops Jinja)
        if linhas_limpas:
            ctx_docx[tag_key + "_BULLETS"] = "\n".join(f"• {l}" for l in linhas_limpas)
        else:
            ctx_docx[tag_key + "_BULLETS"] = "-"

    # Construção de tabelas dinâmicas baseadas nos dados originais
    ctx_docx["DADOS_CADASTRAIS_LIST"] = []
    ident_list = contexto.get("identificacao")
    if ident_list:
        ident_list = ident_list if isinstance(ident_list, list) else [ident_list]
        for d in ident_list:
            ctx_docx["DADOS_CADASTRAIS_LIST"].append({
                "proprietario": limpar_string(d.get("proprietario")),
                "inscricao": str(d.get("inscricao_imobiliaria") or "-") + " / " + str(d.get("numero_cadastral") or "-"),
                "endereco": str(d.get("logradouro", "")) + ", " + str(d.get("numero", "")),
                "loteamento": str(d.get("loteamento", "-")) + " Qd:" + str(d.get("quadra", "-")) + " Lt:" + str(d.get("lote", "-")),
                "area": str(d.get("area_m2") or "-")
            })

    ctx_docx["TABELA_TESTADAS_LIST"] = []
    segmentos = contexto.get("segmentos_limites", [])
    for s in segmentos:
        if isinstance(s, dict):
            ctx_docx["TABELA_TESTADAS_LIST"].append({
                "limite": str(s.get("logradouro") or s.get("tipo_limite") or s.get("confrontante") or "-"),
                "comprimento": "{:.2f}".format(float(s.get("comprimento_m", 0)))
            })
            
    ctx_docx["TABELA_INCLINACAO_LIST"] = []
    incl_dict = contexto.get("inclinacao", {})
    if isinstance(incl_dict, dict):
        faixas = incl_dict.get("faixas", [])
        for f in faixas:
            if isinstance(f, dict):
                cor_hex = str(f.get("cor", "#FFFFFF")).replace("#", "").upper()
                if not cor_hex or cor_hex == "FFFFFF": cor_hex = "F2F2F2"
                ctx_docx["TABELA_INCLINACAO_LIST"].append({
                    "faixa": str(f.get("faixa", "-")),
                    "area": "{:.2f}".format(float(f.get("area_m2", 0))),
                    "perc": "{:.1f}".format(float(f.get("percentual", 0))),
                    "notas": "APP" if bool(f.get("app")) else "-",
                    "cor": cor_hex
                })
                
    # Zonas Multiplas
    ctx_docx["TABELA_ZONAS_LIST"] = []
    zr = contexto.get("zoneamento_resolvido", {})
    zonas_res = zr.get("zonas", [])
    if zonas_res:
        for z in zonas_res:
            param = z.get("parametros", {})
            extras = param.get("extras", {})
            ctx_docx["TABELA_ZONAS_LIST"].append({
                "codigo": str(z.get("codigo") or ""),
                "area": "{:.2f}".format(float(z.get("area_m2", 0))),
                "perc": "{:.1f}".format(float(z.get("percentual_area", 0))),
                "ca_max": str(param.get("CA_max") or "-"),
                "ca_bas": str(param.get("CA_bas") or "-"),
                "tps": str(param.get("Tperm") or "-"),
                "tos": str(param.get("Tocup") or "-"),
                "np_bas": str(param.get("Npav_bas") or "-"),
                "np_max": str(param.get("Npav_max") or "-"),
                "rf": str(extras.get("RF") or "-")
            })
    else:
        z_nome = contexto.get("zoneamento", {}).get("zona", "-")
        ctx_docx["TABELA_ZONAS_LIST"].append({
            "codigo": str(z_nome),
            "area": str(contexto.get("area_lote_m2", "-")),
            "perc": "100.0",
            "ca_max": "-", "ca_bas": "-", "tps": "-", "tos": "-", "np_bas": "-", "np_max": "-", "rf": "-"
        })

    # Risco Geológico
    risco = contexto.get("risco", {})
    classe_inund = str(risco.get("classe_inundacao") or "Não classificado")
    classe_mov = str(risco.get("classe_movimento_massa") or "Não classificado")
    def c_risco(c):
        s = c.upper()
        if "ALTA" in s or "ALTO" in s or s in ("A", "4"): return ("ALTA", "FFCCCC", "Estudo hidrológico / geotécnico obrigatório e elevação/contenção mandatória.")
        if "MÉDIA" in s or "MEDIA" in s or s in ("M", "3"): return ("MÉDIA", "FFEB9C", "Exige investigação geotécnica preliminar.")
        if "BAIXA" in s or "BAIXO" in s or s in ("B", "2"): return ("BAIXA", "CCFFCC", "Padrão construtivo convencional aceito.")
        if "MUITO BAIXA" in s or "MB" in s or s == "1": return ("MUITO BAIXA", "CCFFCC", "Padrão.")
        return ("Não Definido", "F2F2F2", "-")

    gi, ci, ri = c_risco(classe_inund)
    gm, cm, rm = c_risco(classe_mov)

    ctx_docx["RISCO_INUND_CLASSE"] = "Suscetibilidade" if "ALTA" in gi or "MÉDIA" in gi else "Baixo Aconselhado"
    ctx_docx["RISCO_INUND_GRAU"] = gi
    ctx_docx["RISCO_INUND_COR"] = ci
    ctx_docx["RISCO_INUND_RECOM"] = ri
    ctx_docx["RISCO_MOV_CLASSE"] = "Deslizamento e Massa" if "ALTA" in gm or "MÉDIA" in gm else "Baixo Aconselhado"
    ctx_docx["RISCO_MOV_GRAU"] = gm
    ctx_docx["RISCO_MOV_COR"] = cm
    ctx_docx["RISCO_MOV_RECOM"] = rm

    # APPs
    amb = contexto.get("ambiente", {})
    em_nuic = amb.get("em_app_faixa_nuic")
    ctx_docx["APP_FAIXA_STATUS"] = "Presente" if em_nuic else "Ausente"
    ctx_docx["APP_FAIXA_LARGURA"] = str(amb.get("largura_faixa_m") or "-")
    ctx_docx["APP_FAIXA_OBS"] = "; ".join([str(n) for n in amb.get("notas", [])][:2]) if em_nuic else "Sem rio mapeado no lote."
    em_mangue = amb.get("em_app_manguezal")
    ctx_docx["APP_MANGUE_STATUS"] = "Presente" if em_mangue else "Ausente"
    nota_mangue = amb.get("notas", [])[2:] if len(amb.get("notas", [])) > 2 else []
    ctx_docx["APP_MANGUE_OBS"] = "; ".join([str(n) for n in nota_mangue]) if em_mangue and nota_mangue else ("Manguezal detectado." if em_mangue else "Fora de manguezais.")

                
    # Variáveis avulsas comuns (Data, Tipo Análise)
    import datetime
    ctx_docx["TIPO_ANALISE"] = "Gleba Unificada" if contexto.get("area_gleba_unificada") else "Lote"
    ctx_docx["DATA_COMPLETA"] = datetime.datetime.now().strftime("%d/%m/%Y")

    # Aqui podemos passar o restante bruto do contexto para o template,
    # as variáveis que tem "_bruto" (como as testadas como lista)
    return {**contexto, **ctx_docx}

class RenderizadorDOCX:
    """Fachada para geração do relatório DOCX nativo."""

    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.modelo_path = os.path.join(self.base_dir, "modelos", "modelo_relatorio.docx")

    def renderizar_e_salvar(self, contexto: dict, caminho_saida: str) -> tuple:
        """
        Lê o modelo base modelo_relatorio.docx, injeta os dados do contexto
        e salva no caminho de saída.
        """
        if 'docxtpl' not in sys.modules:
            QgsMessageLog.logMessage("Erro: Dependência docxtpl não foi instalada com sucesso.", "Zôni v2", Qgis.Critical)
            return False, "Dependência 'docxtpl' não está instalada ou falhou ao carregar."

        if not os.path.exists(self.modelo_path):
            QgsMessageLog.logMessage(f"Modelo não encontrado: {self.modelo_path}", "Zôni v2", Qgis.Critical)
            return False, f"Arquivo modelo não encontrado em:\n{self.modelo_path}"
            
        try:
            doc = DocxTemplate(self.modelo_path)
            # Pré processamento para converter campos de UI para Jinja do Word
            ctx_pronto = mapear_contexto_html_para_docx(contexto)
            
            doc.render(ctx_pronto)
            doc.save(caminho_saida)
            QgsMessageLog.logMessage(f"Relatório gerado em: {caminho_saida}", "Zôni v2", Qgis.Success)
            return True, ""
        except Exception as e:
            QgsMessageLog.logMessage(f"Erro ao gerar DOCX: {e}", "Zôni v2", Qgis.Critical)
            return False, str(e)

