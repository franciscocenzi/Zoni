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
            # Limpeza rápida e crua de <li class="">Texto</li> para array de strings do Word
            linhas = html_val.replace("</li>", "").split("<li>")
            linhas_limpas = [linha.replace("<b>", "").replace("</b>", "").strip() for linha in linhas if linha.strip()]
            ctx_docx[tag_key + "_ARRAY"] = linhas_limpas
        elif isinstance(html_val, list):
            ctx_docx[tag_key + "_ARRAY"] = html_val
        else:
            ctx_docx[tag_key + "_ARRAY"] = []

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
                ctx_docx["TABELA_INCLINACAO_LIST"].append({
                    "faixa": str(f.get("faixa", "-")),
                    "area": "{:.2f}".format(float(f.get("area_m2", 0))),
                    "perc": "{:.1f}".format(float(f.get("percentual", 0))),
                    "notas": "APP" if bool(f.get("app")) else "-"
                })
                
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

