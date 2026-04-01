import os
import re
import subprocess
import sys
from datetime import datetime

# Adiciona o diretório raiz do plugin ao sys.path para permitir imports absolutos
plugin_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if plugin_dir not in sys.path:
    sys.path.insert(0, plugin_dir)

from infraestrutura.logs.build_logger import log_build

def get_commit_message():
    """Tenta obter a mensagem do commit atual via COMMIT_EDITMSG."""
    try:
        msg_path = os.path.join(plugin_dir, ".git", "COMMIT_EDITMSG")
        if os.path.exists(msg_path):
            with open(msg_path, "r", encoding="utf-8") as f:
                lines = [l.strip() for l in f.readlines() if l.strip() and not l.startswith("#")]
                return lines[0] if lines else ""
    except Exception:
        pass
    return ""

def atualizar_changelog(versao: str, mensagem: str):
    """Adiciona entrada no topo do CHANGELOG.md com a nova versão."""
    changelog_path = os.path.join(plugin_dir, "CHANGELOG.md")
    if not os.path.exists(changelog_path):
        return

    data_hoje = datetime.now().strftime("%Y-%m-%d")
    
    # Classifica a mensagem pelo prefixo convencional
    msg_lower = mensagem.lower()
    if msg_lower.startswith("fix"):
        categoria = "### Corrigido"
    elif msg_lower.startswith("feat"):
        categoria = "### Adicionado"
    elif msg_lower.startswith("refactor"):
        categoria = "### Refatorado"
    elif msg_lower.startswith("perf"):
        categoria = "### Performance"
    elif msg_lower.startswith("docs"):
        categoria = "### Documentação"
    else:
        categoria = "### Alterado"

    nova_entrada = (
        f"\n## [{versao}] — {data_hoje}\n"
        f"{categoria}\n"
        f"- {mensagem}\n"
        f"\n---\n"
    )

    with open(changelog_path, "r", encoding="utf-8") as f:
        conteudo = f.read()

    # Insere após o cabeçalho (primeira linha em branco após o título)
    if "---" in conteudo:
        idx = conteudo.index("---")
        novo_conteudo = conteudo[:idx + 3] + nova_entrada + conteudo[idx + 3:]
    else:
        novo_conteudo = conteudo + nova_entrada

    with open(changelog_path, "w", encoding="utf-8") as f:
        f.write(novo_conteudo)

    subprocess.run(["git", "add", "CHANGELOG.md"], cwd=plugin_dir)

def increment_version():
    metadata_path = os.path.join(plugin_dir, "metadata.txt")
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
            mensagem_commit = get_commit_message()
            if mensagem_commit:
                atualizar_changelog(current_version, mensagem_commit)

if __name__ == "__main__":
    increment_version()
