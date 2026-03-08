# -*- coding: utf-8 -*-
"""Módulo espacial - re-exporta funções principais."""

from .geometrias import unir_geometrias, UtilsGeometria
from .validadores import lotes_sao_contiguos, ValidadorGeometrias

# Você pode adicionar outras exportações conforme necessário
__all__ = [
    'unir_geometrias',
    'UtilsGeometria',
    'lotes_sao_contiguos',
    'ValidadorGeometrias',
]