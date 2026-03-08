import sys
import zipfile
import xml.etree.ElementTree as ET
import json
import os

def get_docx_tables(path):
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = ET.XML(xml_content)

    WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    TABLE = WORD_NAMESPACE + 'tbl'
    ROW = WORD_NAMESPACE + 'tr'
    CELL = WORD_NAMESPACE + 'tc'
    PARA = WORD_NAMESPACE + 'p'
    TEXT = WORD_NAMESPACE + 't'

    tables = []
    for table in tree.iter(TABLE):
        parsed_table = []
        for row in table.iter(ROW):
            parsed_row = []
            for cell in row.iter(CELL):
                paragraphs = []
                for paragraph in cell.iter(PARA):
                    texts = [node.text for node in paragraph.iter(TEXT) if node.text]
                    if texts:
                        paragraphs.append(''.join(texts))
                parsed_row.append('\\n'.join(paragraphs))
            parsed_table.append(parsed_row)
        tables.append(parsed_table)
    return tables

if __name__ == "__main__":
    docx_path = r"c:\Users\franciscocenzi\AppData\Roaming\QGIS\QGIS3\profiles\default\python\plugins\zoni_v2\infraestrutura\dados\2025.0275 - Lei Complementar 275.2025 - Zoneamento - Anexo III.docx"
    tables = get_docx_tables(docx_path)
    # Dump just to a temp JSON so we can analyze it
    out_path = os.path.join(os.path.dirname(docx_path), "extracted_tables.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(tables, f, indent=2, ensure_ascii=False)
    print(f"Extracted {len(tables)} tables to {out_path}")
