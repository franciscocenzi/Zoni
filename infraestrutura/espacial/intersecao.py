# -*- coding: utf-8 -*-
"""Serviço unificado para interseções geométricas."""

from .intersecao_app import intersecao_app, ResultadoAPP
from .intersecao_risco import intersecao_risco, ResultadoRisco
from .intersecao_zoneamento import intersecao_zoneamento, ResultadoZoneamento
from .intersecao_inclinacao import analisar_inclinacao_terreno, ResultadoInclinacao
from .zoneamento_lote import calcular_zoneamento_incidente, ResultadoZoneamentoGeom

class ServicoIntersecao:
    """
    Serviço que encapsula as operações espaciais de interseção do lote 
    com as diversas camadas temáticas (zoneamento, restrições e riscos).
    """

    def __init__(self):
        pass

    def intersecao_zoneamento(self, geom_lote) -> ResultadoZoneamento:
        return intersecao_zoneamento(geom_lote)

    def intersecao_app(self, geom_lote) -> ResultadoAPP:
        return intersecao_app(geom_lote)

    def intersecao_risco(self, geom_lote) -> ResultadoRisco:
        return intersecao_risco(geom_lote)

    def analisar_inclinacao(self, geom_lote, camada_inclinacao, area_lote_m2) -> ResultadoInclinacao:
        if camada_inclinacao is None:
            return None
        return analisar_inclinacao_terreno(geom_lote, camada_inclinacao, area_lote_m2)

    def calcular_zoneamento_incidente(self, geom_lote, camada_zoneamento) -> ResultadoZoneamentoGeom:
        if camada_zoneamento is None:
            return None
        return calcular_zoneamento_incidente(geom_lote, camada_zoneamento)
