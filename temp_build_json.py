import json
import re
import os

def clean_float(val):
    if not val or val.strip() in ['-', '']:
        return None
    # Remove markers like (12)
    val = re.sub(r'\(\d+\)', '', val).strip()
    # Remove %
    val = val.replace('%', '')
    # Handle numbers like 19,95 or 1.000,00
    val = val.replace('.', '').replace(',', '.')
    try:
        return float(val)
    except:
        return None

def extract_notas(text):
    if not text:
        return []
    # Match strings inside parentheses that are mostly digits
    matches = re.findall(r'\((\d{1,2})\)', text)
    return list(dict.fromkeys(matches)) # deduplicate

def parse_split_param(val):
    if not val or val.strip() in ['-', '']:
        return None, None
    val = re.sub(r'\(\d+\)', '', val).strip()
    if '/' in val:
        parts = val.split('/')
        return clean_float(parts[0]), clean_float(parts[1])
    else:
        # e.g. "Livre/Livre" or just a single number
        if "Livre" in val:
            return 999, 999.0
        return clean_float(val), None

def parse_docx_json(input_json, out_params_path, out_notas_zona_path):
    with open(input_json, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    table = data[0]
    # Skip header headers
    rows = table[3:]
    
    params_dict = {}
    notas_zona = {}
    
    for row in rows:
        if len(row) < 14:
            continue
            
        zona_raw = row[0]
        if not zona_raw.strip():
            continue
            
        # Try finding the code at the end (before parentheses)
        parts = zona_raw.split('–')
        if len(parts) > 1:
            code = parts[-1].split('(')[0].strip()
        else:
            code = zona_raw.split('(')[0].strip()
                
        # Fix aliases or known codes
        if code == "EU 1": code = "EU1"
        if code == "EU 2": code = "EU2"
        if code == "EU 3": code = "EU3"
        if code == "EU 4": code = "EU4"
        if code == "MUQ 1": code = "MUQ1"
        if code == "MUQ 2": code = "MUQ2"
        if code == "MUQ 3": code = "MUQ3"
        if code == "MUQ 4": code = "MUQ4"
        if code == "MUQ 5": code = "MUQ5"
        if code == "MUQ 6": code = "MUQ6"
        if code == "MUCON 1": code = "MUCON1"
        if code == "MUCON 2": code = "MUCON2"
        if code == "MUCON 3": code = "MUCON3"
        if code == "MUCON 4": code = "MUCON4"
        if code == "MUCON 5": code = "MUCON5"
        if code == "MUCON 6": code = "MUCON6"
        if code == "MUCON 7": code = "MUCON7"
        
        # Accumulate all notes from all columns for this row
        all_notas = []
        for cell in row:
            all_notas.extend(extract_notas(cell))
            
        # Deduplicate and sort
        all_notas = sorted(list(set(all_notas)), key=lambda x: int(x))
        notas_zona[code] = all_notas
        
        ca_min = clean_float(row[1])
        ca_bas = clean_float(row[2])
        ca_max = clean_float(row[3])
        
        tps = clean_float(row[4])
        if tps and tps > 1.0: tps = tps / 100.0 # Convert 5% to 0.05 if needed
        tos = clean_float(row[5])
        if tos and tos > 1.0: tos = tos / 100.0
            
        rf = clean_float(row[6])
        rlf, _ = parse_split_param(row[7]) # 0.00/0.00 -> 0.0
        
        np_bas, gab_bas = parse_split_param(row[8])
        np_max, gab_max = parse_split_param(row[9])
        
        hemb = clean_float(row[10])
        vagas = 1 # Fallback, lei exige 1 por padrao

        # Build json object matching ParametrosZona pattern
        p = {
            "indices": {
                "CA_min": ca_min,
                "CA_bas": ca_bas,
                "CA_max": ca_max,
                "Tperm": tps,
                "Tocup": tos,
                "Npav_bas": int(np_bas) if np_bas else None,
                "Npav_max": int(np_max) if np_max else None,
                "Gab_bas": gab_bas,
                "Gab_max": gab_max,
                "RF": rf,
                "RLF": rlf,
                "HEMB": hemb,
                "vagas_min": vagas
            }
        }
        params_dict[code] = p

    # Read txt file for the text of the notes
    txt_path = os.path.join(os.path.dirname(out_params_path), "2025.0275 - Lei Complementar 275.2025 - Zoneamento - Anexo III - Notas.txt")
    notas_explicadas = {}
    if os.path.exists(txt_path):
        with open(txt_path, 'r', encoding='utf-8') as f:
            for line in f:
                match = re.match(r'^\((\d+)\)\s*(.*)', line.strip())
                if match:
                    n_id = match.group(1).zfill(2) # e.g. "  5" -> "05" or just keep as is. Let's keep as is matching the array
                    n_id = match.group(1) 
                    text = match.group(2).strip()
                    notas_explicadas[n_id] = text

    with open(out_params_path, 'w', encoding='utf-8') as f:
        json.dump(params_dict, f, indent=2, ensure_ascii=False)
        
    with open(out_notas_zona_path, 'w', encoding='utf-8') as f:
        json.dump(notas_zona, f, indent=2, ensure_ascii=False)
        
    out_explicadas_path = os.path.join(os.path.dirname(out_params_path), "notas_explicadas.json")
    with open(out_explicadas_path, 'w', encoding='utf-8') as f:
        json.dump(notas_explicadas, f, indent=2, ensure_ascii=False)

    print("Success! Generated parameters and notes.")

if __name__ == "__main__":
    base = r"c:\Users\franciscocenzi\AppData\Roaming\QGIS\QGIS3\profiles\default\python\plugins\zoni_v2\infraestrutura\dados"
    in_json = os.path.join(base, "extracted_tables.json")
    out_param = os.path.join(base, "parametros_zoneamento.json")
    out_notas = os.path.join(base, "zoni_notas_por_zona.json")
    parse_docx_json(in_json, out_param, out_notas)
