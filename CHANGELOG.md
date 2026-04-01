# Changelog — Zôni v2

Todas as mudanças relevantes do plugin são documentadas aqui.  
Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/).

---
## [2.0.1.010] — 2026-03-31
### Corrigido
- fix: Corrigir repeticao de linhas em tabela (uso de %tr for correto), larguras fixas de coluna e estilo visual aprimorado

---

## [2.0.1.009] — 2026-03-31
### Documentação
- docs: Adicionar CHANGELOG.md retroativo e auto-atualizacao no pre-commit via increment_version.py

---

## [2.0.1.008] — 2026-03-31
### Corrigido
- fix: Corrigir error de Jinja - cellbg nao pode ser envolvido em {{ }}, notas convertidas para strings pre-montadas com bullets

---


## [2.0.1.007] — 2026-03-31
### Corrigido
- Erro de parsing Jinja `Expected an expression, got 'end of print statement'`: tags `{% cellbg %}` não podem ser embrulhadas em `{{ }}` — corrigido com detecção automática de tags de bloco na função geradora de tabelas DOCX.
- Notas e condicionantes não podiam usar `{% for %}` dentro de parágrafos (causa XML inválido no Word): convertidas para strings pré-montadas com bullet points `•` no `renderizador_docx.py`.

---

## [2.0.1.006] — 2026-03-31
### Adicionado
- Popup `QMessageBox.warning` exibido na tela quando nenhum lote está selecionado ao clicar em Analisar (substituiu o aviso silencioso na barra do QGIS).
- Tabelas de **Risco Geoambiental** (Inundação e Movimentos de Massa) com células coloridas (verde/amarelo/vermelho) e bordas invisíveis (`Normal Table`).
- Tabela de **Inclinação do Terreno** com coluna de legenda colorida pelo HEX fornecido pelo motor analítico.
- Suporte a **Múltiplas Zonas** na tabela de Zoneamento (dinâmica, iterando sobre `TABELA_ZONAS_LIST`).
- Seção "7. Notas e Condicionantes Técnicas" com listas separadas por categoria (Anexos, Condicionantes, Restrições).
- Variáveis `APP_FAIXA_STATUS`, `APP_MANGUE_STATUS` e observações injetadas no template.

### Corrigido
- A função `adicionar_tabela_jinja` aceitava apenas campos simples; estendida para detectar e passar tags de bloco `{% %}` sem envolve-las erroneamente.

---

## [2.0.1.005] — 2026-03-31
### Corrigido
- Tags `{% tr for %}` e `{% tr endfor %}` obsoletas removidas do `document.xml` do modelo DOCX — causavam `Encountered unknown tag 'tr'` no docxtpl moderno (tags corretas são `{% for %}` / `{% endfor %}`).
- Script `gerar_template_base_docx.py` atualizado para gerar as tags corretas desde o início.

---

## [2.0.1.004] — 2026-03-31
### Corrigido
- Erro `'str' object has no attribute 'get'`: o mapeador de contexto estava iterando sobre strings em vez de dicionários nas chaves `inclinacao` e `segmentos_limites`.
- Chaves de contexto corrigidas: `inclinacao.faixas`, `segmentos_limites[].logradouro`, `segmentos_limites[].comprimento_m`.
- Adicionadas variáveis automáticas `DATA_COMPLETA` e `TIPO_ANALISE` ao contexto DOCX.

---

## [2.0.1.003] — 2026-03-31
### Corrigido
- Plugin abria nova instância do QGIS ao instalar `docxtpl` via pip: corrigido usando `sys.prefix/python.exe` em vez de `sys.executable` (que retornava `qgis-bin.exe`).

### Adicionado
- Feedback explícito de erro via `QMessageBox` quando a geração do DOCX falha.
- Memória do último diretório de exportação utilizando `QgsSettings`.

---

## [2.0.1.001] — 2026-03-31
### Adicionado
- Motor de geração de relatórios **DOCX nativo** via `docxtpl` (substitui o HTML+WebView).
- Auto-instalador silencioso do `docxtpl` via pip ao carregar o plugin.
- Template base `modelo_relatorio.docx` gerado pelo script `gerar_template_base_docx.py`.
- Script `scripts/increment_version.py` + hook `pre-commit` para incremento automático de versão a cada commit.
- Abertura automática do arquivo Word gerado após salvar (`os.startfile`).

### Removido
- Visualização via WebView HTML (substituída pelo fluxo nativo DOCX).

---

## [2.0.0.015] — 2026-03-08
### Estado
- Versão base da migração: motor HTML funcional com templates completos, formatação PT-BR, dados de APP, risco, inclinação e zoneamento resolvido.

### Problemas Conhecidos (na época)
- Exportação resultava em arquivo de texto puro (não havia conversão real para PDF/DOCX).
- Layout visual dependia de CSS que não era renderizado fora do browser embutido do QGIS.

---

## Problemas Conhecidos (versão atual 2.0.1.007)
- `{% cellbg %}` na tabela de inclinação ainda não foi testado em produção — pode exigir ajuste se o Word reportar erro de macro.
- A seção de Mapa (`{{ IMAGEM_MAPA }}`) ainda não injeta imagem real; exibe texto vazio.
- Algumas instalações do QGIS podem não ter o `python.exe` em `sys.prefix` — o instalador silencioso do `docxtpl` pode falhar silenciosamente nesses casos.
