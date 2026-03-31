# -*- coding: utf-8 -*-
"""Container de dependências para injeção."""

import os
from dataclasses import dataclass
from typing import Optional
from qgis.core import QgsProject

from ..dominio.motores.motor_analise_lote import MotorAnaliseLote
from ..dominio.regras.regras_zoneamento import RegrasZoneamento
from ..dominio.regras.regras_app import RegrasAPP
from ..dominio.regras.regras_risco import RegrasRisco
from ..infraestrutura.espacial.geometrias import UtilsGeometria
from ..infraestrutura.espacial.intersecao import ServicoIntersecao
from ..infraestrutura.espacial.testadas import ServicoTestadas
from ..infraestrutura.espacial.validadores import ValidadorGeometrias
from ..infraestrutura.relatorios.construtor_relatorio import ConstrutorRelatorio
from ..infraestrutura.relatorios.renderizador_docx import RenderizadorDOCX


@dataclass
class Config:
    """Configurações do sistema."""
    caminho_parametros: Optional[str] = None
    max_dist_testada_m: float = 20.0


class Container:
    """Container principal de dependências."""
    
    def __init__(self):
        # Configurações
        self.config = Config()
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.config.caminho_parametros = os.path.join(
            base_dir, "infraestrutura", "dados", "parametros_urbanisticos.json"
        )
        
        # Serviços de infraestrutura
        self.utils_geometria = UtilsGeometria()
        self.servico_intersecao = ServicoIntersecao()
        self.servico_testadas = ServicoTestadas()
        self.validador = ValidadorGeometrias()
        
        # Regras de domínio
        self.regras_zoneamento = RegrasZoneamento()
        self.regras_app = RegrasAPP()
        self.regras_risco = RegrasRisco()
        
        # Motor de análise
        self.motor_analise = MotorAnaliseLote(
            regras_zoneamento=self.regras_zoneamento,
            regras_app=self.regras_app,
            regras_risco=self.regras_risco,
            utils_geometria=self.utils_geometria,
            servico_intersecao=self.servico_intersecao,
            servico_testadas=self.servico_testadas,
            validador=self.validador
        )
        
        # Relatórios
        self.construtor_relatorio = ConstrutorRelatorio()
        self.renderizador_docx = RenderizadorDOCX()
        
        # Estado
        self.projeto = QgsProject.instance()
        
    def obter_camada(self, nome_camada: str):
        """Obtém uma camada pelo nome."""
        return self.projeto.mapLayersByName(nome_camada)[0] if self.projeto.mapLayersByName(nome_camada) else None
