# -*- coding: utf-8 -*-
"""Regras de zoneamento urbano – LC 275/2025 (Versão Robusta)."""

from dataclasses import dataclass, field
from typing import Dict, Optional, Any, List
import json

@dataclass
class ParametrosZona:
    codigo: str
    CA_min: Optional[float] = None
    CA_bas: Optional[float] = None
    CA_max: Optional[float] = None
    Tperm: Optional[float] = None
    Tocup: Optional[float] = None
    Npav_bas: Optional[int] = None
    Npav_max: Optional[int] = None
    Gab_bas: Optional[float] = None
    Gab_max: Optional[float] = None
    # Lista de códigos de notas/condicionantes incidentes na zona
    notas: List[str] = field(default_factory=list)
    extras: Dict[str, Any] = field(default_factory=dict)

@dataclass
class ResultadoAvaliacaoZona:
    zona: str
    parametros: ParametrosZona
    conforme: bool
    pendencias: List[str] = field(default_factory=list)
    observacoes: List[str] = field(default_factory=list)
    valores_calculados: Dict[str, Any] = field(default_factory=dict)

def _limpar_float(valor: Any) -> Optional[float]:
    """Converte valores variados (str com vírgula, int, float) para float robusto."""
    if valor in (None, "", "NULL", "null", " "): return None
    if isinstance(valor, (int, float)): return float(valor)
    try:
        # Lida com padrão brasileiro: "1.250,50" -> "1250.50"
        s = str(valor).strip().replace(".", "").replace(",", ".")
        return float(s)
    except (ValueError, TypeError):
        return None

def carregar_parametros_de_arquivo(caminho_json: str) -> Dict[str, ParametrosZona]:
    """Carrega o JSON e garante que todos os índices sejam floats ou ints válidos."""
    try:
        with open(caminho_json, "r", encoding="utf-8") as f:
            bruto = json.load(f)
    except Exception as e:
        print(f"Erro crítico ao ler arquivo JSON: {e}")
        return {}

    parametros_por_zona = {}
    for codigo, dados in bruto.items():
        indices = dados.get("indices", {}) or {}
        notas_raw = dados.get("notes", []) or []
        notas_limpa: List[str] = []
        for n in notas_raw:
            try:
                notas_limpa.append(str(int(n)))
            except (TypeError, ValueError):
                if n is not None:
                    notas_limpa.append(str(n))
        
        # Mapeamento com limpeza automática de tipos
        p = ParametrosZona(
            codigo=codigo,
            CA_min=_limpar_float(indices.get("CA_min")),
            CA_bas=_limpar_float(indices.get("CA_bas")),
            CA_max=_limpar_float(indices.get("CA_max")),
            Tperm=_limpar_float(indices.get("Tperm")),
            Tocup=_limpar_float(indices.get("Tocup")),
            Npav_bas=indices.get("Npav_bas"), # Int
            Npav_max=indices.get("Npav_max"), # Int
            Gab_bas=_limpar_float(indices.get("Gab_bas")),
            Gab_max=_limpar_float(indices.get("Gab_max")),
            notas=notas_limpa,
            extras={k: v for k, v in indices.items() if k not in ParametrosZona.__annotations__}
        )
        parametros_por_zona[codigo.upper().replace(" ", "")] = p

    return parametros_por_zona

def avaliar_edificacao_na_zona(
    zona: str,
    parametros: ParametrosZona,
    area_lote_m2: float,
    **kwargs # Permite receber area_construida, area_ocupada, etc.
) -> ResultadoAvaliacaoZona:
    """Avalia o projeto contra os parâmetros da zona com tolerância de precisão."""
    
    pendencias, observacoes, valores = [], [], {}
    eps = 1e-6 # Margem de erro para cálculos de ponto flutuante

    if area_lote_m2 <= 0:
        raise ValueError("Área do lote inválida para cálculo.")

    # 1. Avaliação de Coeficiente de Aproveitamento (CA)
    area_const = kwargs.get("area_construida_total_m2")
    if area_const is not None:
        ca_real = area_const / area_lote_m2
        valores["CA_real"] = ca_real
        if parametros.CA_min and ca_real < (parametros.CA_min - eps):
            pendencias.append(f"CA real ({ca_real:.2f}) abaixo do mínimo ({parametros.CA_min:.2f}).")
        if parametros.CA_max and ca_real > (parametros.CA_max + eps):
            pendencias.append(f"CA real ({ca_real:.2f}) acima do máximo ({parametros.CA_max:.2f}).")
    else:
        observacoes.append("CA não avaliado: Área construída não informada.")

    # 2. Avaliação de Taxa de Ocupação (TO) - Converte decimal (0.6) para % (60%)
    area_ocup = kwargs.get("area_ocupada_projecao_m2")
    if area_ocup is not None:
        to_real = (area_ocup / area_lote_m2) * 100
        valores["Tocup_real"] = to_real
        if parametros.Tocup and to_real > (parametros.Tocup + eps):
            pendencias.append(f"TO real ({to_real:.1f}%) acima do máximo ({parametros.Tocup:.1f}%).")
    else:
        observacoes.append("TO não avaliada.")

    # 3. Avaliação de Pavimentos e Gabarito
    n_pav = kwargs.get("numero_pavimentos")
    if n_pav and parametros.Npav_max and n_pav > parametros.Npav_max:
        pendencias.append(f"Pavimentos ({n_pav}) excedem o máximo ({parametros.Npav_max}).")

    alt_max = kwargs.get("altura_maxima_m")
    if alt_max and parametros.Gab_max and alt_max > (parametros.Gab_max + eps):
        pendencias.append(f"Altura ({alt_max}m) excede o gabarito ({parametros.Gab_max}m).")

    return ResultadoAvaliacaoZona(
        zona=zona,
        parametros=parametros,
        conforme=len(pendencias) == 0,
        pendencias=pendencias,
        observacoes=observacoes,
        valores_calculados=valores
    )
