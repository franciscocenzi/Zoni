import os
import re
import subprocess
import sys

# Adiciona o diretório raiz do plugin ao sys.path para permitir imports absolutos
plugin_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if plugin_dir not in sys.path:
    sys.path.insert(0, plugin_dir)

from infraestrutura.logs.build_logger import log_build

def increment_version():
    metadata_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "metadata.txt")
    if not os.path.exists(metadata_path):
        print(f"Arquivo metadata.txt não encontrado: {metadata_path}")
        return
        
    with open(metadata_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    def replace_version(match):
        prefix = match.group(1)
        minor_str = match.group(2)
        minor_int = int(minor_str) + 1
        # preserve padding
        new_minor = str(minor_int).zfill(len(minor_str))
        print(f"Zôni v2: Versão incrementada de {prefix}{minor_str} para {prefix}{new_minor}")
        return f"version={prefix}{new_minor}"
        
    new_content, count = re.subn(r'^version=(\d+\.\d+\.\d+\.)(\d+)$', replace_version, content, flags=re.MULTILINE)
    # Helper to extract the version string after replacement
    def get_current_version(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.startswith('version='):
                        return line.strip().split('=')[1]
        except Exception:
            return "unknown"
        return "unknown"
    
    if count > 0:
        with open(metadata_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        # Add metadata.txt to the current commit
        subprocess.run(["git", "add", "metadata.txt"], cwd=os.path.dirname(metadata_path))
        
        # Log the new build version
        current_version = get_current_version(metadata_path)
        if current_version != "unknown":
            log_build(current_version, "incrementado via pre‑commit")
if __name__ == "__main__":
    increment_version()
