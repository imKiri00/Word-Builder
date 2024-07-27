# placeholder_replacer.py
import re

def replace_placeholders(doc, data):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            replaced_text = re.sub(r'{{([A-Z_]+)}}', lambda m: data.get(m.group(1), m.group(0)), run.text)
            run.text = replaced_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        replaced_text = re.sub(r'{{([A-Z_]+)}}', lambda m: data.get(m.group(1), m.group(0)), run.text)
                        run.text = replaced_text

def apply_replacements(doc, data):
    replace_placeholders(doc, data)
    return doc