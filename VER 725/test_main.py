# main.py
from html_to_docx_converter import HTMLToDocxConverter
from placeholder_replacer import apply_replacements
import json

def generate_document(html_template, replacement_data, output_filename):
    # Convert HTML to DOCX
    converter = HTMLToDocxConverter()
    doc = converter.convert(html_template)
    
    # Replace placeholders with data
    doc = apply_replacements(doc, replacement_data)
    
    # Save the final document
    doc.save(output_filename)
    print(f"Document saved as {output_filename}")

if __name__ == "__main__":
    # Load HTML template
    html_template = """
    
    """
    # Load replacement data
    replacement_data = json.loads('''
    {
    "COUNTRY": "РЕПУБЛИКА СРБИЈА"
    }
    ''')
    
    # Generate the document
    generate_document(html_template, replacement_data, 'output_document.docx')
