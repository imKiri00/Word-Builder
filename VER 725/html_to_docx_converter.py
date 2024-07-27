# html_to_docx_converter.py

from docx import Document
from bs4 import BeautifulSoup
from element_processors import (
    process_heading,
    process_paragraph,
    process_table,
    process_list,
    process_div
)
from config import MARGINS

class HTMLToDocxConverter:
    def __init__(self):
        self.doc = Document()
        self._set_margins()

    def _set_margins(self):
        sections = self.doc.sections
        for section in sections:
            section.top_margin = MARGINS['top']
            section.bottom_margin = MARGINS['bottom']
            section.left_margin = MARGINS['left']
            section.right_margin = MARGINS['right']

    def convert(self, html_content):
        soup = BeautifulSoup(html_content, 'html.parser')
        
        for element in soup.find_all(recursive=False):
            if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                process_heading(self.doc, element)
            elif element.name == 'p':
                process_paragraph(self.doc, element)
            elif element.name == 'table':
                process_table(self.doc, element)
            elif element.name == 'ul':
                process_list(self.doc, element)
            elif element.name == 'div':
                process_div(self.doc, element)

        return self.doc