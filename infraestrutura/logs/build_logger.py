import os
import datetime

# Diretório onde o log será armazenado (criado ao importar este módulo)
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

LOG_FILE = os.path.join(LOG_DIR, "build.log")

def _write_entry(entry: str):
    """Escreve uma linha de log com timestamp no arquivo de log."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} - {entry}\n")

def log_build(version: str, description: str = ""):
    """Registra uma nova build.
    Args:
        version: string da versão, ex.: "2.0.0.011"
        description: texto livre opcional.
    """
    entry = f"BUILD {version}" + (f" – {description}" if description else "")
    _write_entry(entry)

def log_report(version: str, parameters: dict, notes: list = None):
    """Registra a geração de um relatório.
    Args:
        version: versão do plugin usada no relatório.
        parameters: dicionário de parâmetros urbanísticos que foram renderizados.
        notes: lista opcional de notas/observações incluídas.
    """
    notes_str = ", ".join(notes) if notes else ""
    entry = f"REPORT version={version} params={list(parameters.keys())}" + (f" notes={notes_str}" if notes_str else "")
    _write_entry(entry)
