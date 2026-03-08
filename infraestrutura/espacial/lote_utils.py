# -*- coding: utf-8 -*-
import unicodedata
from typing import Dict, Any, List, Optional
from qgis.core import QgsFeature, QgsGeometry

MAPEAMENTO_CAMPOS_LOTE = {
    "inscricao": ["inscr_imob", "inscricao", "INSCRICAO", "NR_CADASTRO", "nr_cadastr", "N_CADASTRO"],
    "matricula": ["Matrícula", "matricula", "MATRICULA", "N_MATRICULA", "MATR"],
    "proprietario": ["Propriet.", "proprietario", "PROPRIETARIO", "NOME_PROPRIETARIO", "DONO", "NOME"],
    "bairro": ["Bairro", "BAIRRO", "nome_bairro", "NOME_BAIRRO"],
    "logradouro": ["Logradouro", "rua", "LOGRADOURO", "NOME_LOGRADOURO", "VIA"],
    "numero": ["numero", "NUMERO", "N_IMOVEL", "n_porta"],
    "area_campo": ["área", "area", "AREA", "area_m2", "AREA_M2"]
}

def normalizar_texto(texto: str) -> str:
    if not texto: return ""
    texto = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8')
    return texto.lower().strip().replace(' ', '_').replace('.', '').replace('º', '')

def buscar_valor_campo_robusto(feicao: QgsFeature, candidatos: List[str]) -> Any:
    """Versão robusta que verifica nulidade, strings vazias e normalização."""
    nomes_reais = feicao.fields().names()
    mapa_norm = {normalizar_texto(n): n for n in nomes_reais}
    
    for c in candidatos:
        # Tenta nome exato
        if c in nomes_reais:
            val = feicao[c]
            if val not in (None, "", "NULL", "NULL", " "): return val
        
        # Tenta nome normalizado
        c_norm = normalizar_texto(c)
        if c_norm in mapa_norm:
            val = feicao[mapa_norm[c_norm]]
            if val not in (None, "", "NULL", "NULL", " "): return val
    return None

def extrair_dados_cadastrais(feicao: QgsFeature) -> Dict[str, Any]:
    """
    Extrai todos os atributos brutos para que o construtor de relatórios 
    possa aplicar sua própria lógica robusta de identificação.
    Adiciona também propriedades úteis como area geométrica e id.
    """
    dados = {}
    for fld in feicao.fields():
        nome = fld.name()
        val = feicao[nome]
        if val in (None, "", "NULL", "NULL", " "):
            val = None
        dados[nome] = val

    dados["id"] = feicao.id()
    
    geom = feicao.geometry()
    dados["area_geom"] = geom.area() if geom and not geom.isEmpty() else 0.0
    return dados
