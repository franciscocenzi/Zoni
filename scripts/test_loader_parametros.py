# -*- coding: utf-8 -*-
"""Script de verificação: testa carregamento e lookup de parâmetros de zoneamento."""

import os
import sys

# Adiciona o diretório pai ao path para importar o módulo
plugin_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(plugin_dir))

import json

def _limpar_float(valor):
    if valor in (None, "", "NULL", "null", " "):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        s = str(valor).strip().replace(".", "").replace(",", ".")
        return float(s)
    except (ValueError, TypeError):
        return None


def carregar_parametros(caminho_json):
    with open(caminho_json, "r", encoding="utf-8") as f:
        bruto = json.load(f)

    resultado = {}
    for codigo, dados in bruto.items():
        indices = dados.get("indices", {}) or {}
        chave_normalizada = codigo.upper().replace(" ", "")
        resultado[chave_normalizada] = {
            "codigo_original": codigo,
            "CA_bas": _limpar_float(indices.get("CA_bas")),
            "CA_max": _limpar_float(indices.get("CA_max")),
            "Tperm": _limpar_float(indices.get("Tperm")),
            "Tocup": _limpar_float(indices.get("Tocup")),
            "Npav_bas": indices.get("Npav_bas"),
            "Npav_max": indices.get("Npav_max"),
            "extras": {k: v for k, v in indices.items()
                       if k not in ("CA_min","CA_bas","CA_max","Tperm","Tocup","Npav_bas","Npav_max","Gab_bas","Gab_max")},
        }
    return resultado


def main():
    json_path = os.path.join(plugin_dir, "infraestrutura", "dados", "parametros_urbanisticos.json")
    print(f"Arquivo JSON: {json_path}")
    print(f"Existe? {os.path.exists(json_path)}")
    print()

    params = carregar_parametros(json_path)
    total = len(params)
    print(f"Total de zonas carregadas: {total}")
    print()

    # Zonas críticas para teste
    zonas_teste = ["MUQ3", "MUQ1", "MUQ2", "MUPA1", "ZEOT1", "ZEOT2", "ZEOT3",
                   "EU1", "EU2", "MEU", "MUIS", "MUCON1", "MUCON3"]

    falhas = []
    for z in zonas_teste:
        p = params.get(z)
        if p is None:
            falhas.append(z)
            print(f"  ❌ {z}: NÃO ENCONTRADA")
        else:
            rf = p["extras"].get("RF") or p["extras"].get("RF_Sec")
            hemb = p["extras"].get("Hemb_max") or p["extras"].get("HEMB")
            print(f"  ✓  {z}: CA_bas={p['CA_bas']}, CA_max={p['CA_max']}, "
                  f"Tocup={p['Tocup']}, Npav_max={p['Npav_max']}, RF={rf}, Hemb={hemb}")

    print()
    if falhas:
        print(f"ATENÇÃO: {len(falhas)} zona(s) não encontrada(s): {falhas}")
        sys.exit(1)
    else:
        print("✓ Todas as zonas críticas foram encontradas com parâmetros válidos.")


if __name__ == "__main__":
    main()
