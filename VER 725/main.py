import sys
from html_to_docx_converter import HTMLToDocxConverter
from placeholder_replacer import apply_replacements
import json

def generate_document(input_html_path, output_docx_path):
    # Read the HTML content
    with open(input_html_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    # Convert HTML to DOCX
    converter = HTMLToDocxConverter()
    doc = converter.convert(html_content)
    
    
     # Load replacement data
    replacement_data = json.loads('''
    {
    "COUNTRY": "РЕПУБЛИКА СРБИЈА"
    }
    ''')
    
    doc = apply_replacements(doc, replacement_data)
    
    # Save the final document
    doc.save(output_docx_path)
    print(f"Document saved as {output_docx_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python main.py <input_html_file> <output_docx_file>")
        sys.exit(1)

    input_html_path = sys.argv[1]
    output_docx_path = sys.argv[2]

    generate_document(input_html_path, output_docx_path)
