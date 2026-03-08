# -*- coding: utf-8 -*-
"""Regras de domínio referentes a riscos geoambientais."""

from ...infraestrutura.espacial.intersecao_risco import ResultadoRisco

class RegrasRisco:
    """Implementa validações e regras sobre riscos geoambientais no lote."""

    def __init__(self):
        pass

    def aplicar(self, res_risco: ResultadoRisco) -> ResultadoRisco:
        """
        Recebe o resultado bruto de interseção de riscos e aplica
        as regras de restrição ou formatação específicas do município.
        
        Atualmente, atua como pass-through com formatação base extensível, 
        permitindo futura expansão modular sem quebrar a análise.
        """
        if not res_risco:
            return res_risco
            
        # Placeholder para regras futuras: se risco alto -> exigir laudo geotécnico.
        
        return res_risco
