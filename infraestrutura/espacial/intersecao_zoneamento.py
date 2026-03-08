# -*- coding: utf-8 -*-
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from qgis.core import QgsFeatureRequest, QgsSpatialIndex, QgsGeometry
from .config_camadas import obter_camada
from .lote_utils import buscar_valor_campo_robusto

@dataclass
class ZonaIncidente:
    codigo: str
    area_m2: float
    percentual: float
    macrozona: Optional[str] = None
    eixos: List[str] = field(default_factory=list)
    especiais: List[str] = field(default_factory=list)

@dataclass
class ResultadoZoneamento:
    zona_principal: Optional[str] = None
    macrozona_principal: Optional[str] = None
    area_total_lote_m2: float = 0.0
    detalhes_zonas: List[ZonaIncidente] = field(default_factory=list)
    mensagens: List[str] = field(default_factory=list)

def intersecao_zoneamento(geom_lote: QgsGeometry) -> ResultadoZoneamento:
    resultado = ResultadoZoneamento()
    if not geom_lote or geom_lote.isEmpty():
        resultado.mensagens.append("Geometria inválida.")
        return resultado

    resultado.area_total_lote_m2 = geom_lote.area()
    camada_zon = obter_camada("zoneamento")
    if not camada_zon:
        resultado.mensagens.append("Camada de zoneamento não encontrada.")
        return resultado

    idx = QgsSpatialIndex(camada_zon.getFeatures())
    ids = idx.intersects(geom_lote.boundingBox())
    mapa_zonas = {}

    for feicao in camada_zon.getFeatures(QgsFeatureRequest().setFilterFids(ids)):
        if not feicao.geometry().intersects(geom_lote): continue
        
        inter = feicao.geometry().intersection(geom_lote)
        if inter.isEmpty(): continue
        
        # Extração Robusta (usando candidatos do original)
        cod = buscar_valor_campo_robusto(feicao, ["ZONEAMENTO", "ZONA", "cod_zona", "SIGLA_ZONA"])
        if not cod: continue
        cod = str(cod).strip().upper()

        if cod not in mapa_zonas:
            mapa_zonas[cod] = {
                "area": 0.0, 
                "macro": buscar_valor_campo_robusto(feicao, ["MACROZONA", "MACRO"]),
                "eixos": [], "especiais": []
            }
        
        mapa_zonas[cod]["area"] += inter.area()

        # Lógica Robusta de Múltiplos Eixos/Especiais (Recuperada)
        for campo, chave in [("eixos", ["EIXO", "EIXOS"]), ("especiais", ["ESPECIAL", "zona_especial"])]:
            val = buscar_valor_campo_robusto(feicao, chave)
            if val:
                # Se for string com delimitadores (ex: "EIXO 1; EIXO 2"), separa em lista
                partes = [p.strip() for p in str(val).replace(",", ";").split(";") if p.strip()]
                mapa_zonas[cod][campo].extend(partes)

    for cod, d in mapa_zonas.items():
        perc = (d["area"] / resultado.area_total_lote_m2 * 100) if resultado.area_total_lote_m2 > 0 else 0
        resultado.detalhes_zonas.append(ZonaIncidente(
            codigo=cod, area_m2=d["area"], percentual=perc,
            macrozona=d["macro"], eixos=list(set(d["eixos"])), especiais=list(set(d["especiais"]))
        ))

    resultado.detalhes_zonas.sort(key=lambda x: x.area_m2, reverse=True)
    if resultado.detalhes_zonas:
        resultado.zona_principal = resultado.detalhes_zonas[0].codigo
        resultado.macrozona_principal = resultado.detalhes_zonas[0].macrozona

    return resultado