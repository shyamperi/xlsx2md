"""Excel2Markdown"""
import zipfile
import xml.etree.ElementTree as ET
import sys
import os


class ExcelToMarkdownConverter:
    """Class to handle conversion of Excel to Markdown using XML parsing."""

    ns = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'relationship': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'drawings_main': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'drawings_sd': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    }

    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.shared_strings = []
        self.styles = {}

    def extract_merged_cells(self, sheet_root):
        """Extract merged cells ranges from sheet XML."""
        merged_cells = sheet_root.find('main:mergeCells', self.ns)
        if merged_cells is None:
            return []

        return [
            merge_cell.attrib['ref']
            for merge_cell in merged_cells.findall('main:mergeCell', self.ns)
        ]

    def extract_shared_strings(self, zip_ref):
        """
        Extract shared strings from sharedStrings.xml
        (if present) with both simple and rich text formatting.
        """
        try:
            with zip_ref.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                self.shared_strings = [
                    ''.join(
                        t.text or ''
                            for t in si.findall('.//main:t', self.ns)
                        ) or si.find('main:t', self.ns).text or ''
                    for si in root.findall('main:si', self.ns)
                ]
        except KeyError:
            pass

    def extract_styles(self, zip_ref):
        """
        Extract styles from styles.xml,
        including font information (superscript, subscript, and indentation).
        """
        try:
            with zip_ref.open('xl/styles.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()

                # Extract font information for superscript/subscript
                for idx, font in enumerate(root.findall('main:fonts/main:font', self.ns)):
                    vert_align = font.find('main:vertAlign', self.ns)
                    if vert_align is not None:
                        self.styles[idx] = {'vertAlign': vert_align.attrib['val']}

                # Extract cell formatting (e.g., indentation)
                for idx, xf in enumerate(root.findall('main:cellXfs/main:xf', self.ns)):
                    alignment = xf.find('main:alignment', self.ns)
                    if alignment is not None and 'indent' in alignment.attrib:
                        self.styles.setdefault(idx, {})['indent'] = int(alignment.attrib['indent'])
        except KeyError:
            pass

    def extract_relationships(self, zip_ref, worksheet_path):
        """Extract relationships from the worksheet relationship file."""
        relationships = {}
        rels_path = worksheet_path.replace('worksheets/sheet', 'worksheets/_rels/sheet') + '.rels'
        try:
            with zip_ref.open(rels_path) as f:
                tree = ET.parse(f)
                relationships = {
                    rel.attrib['Id']: rel.attrib['Target']
                    for rel in tree.findall('relationship:Relationship', self.ns)
                }
        except KeyError:
            pass
        return relationships

    def extract_full_text(self, root, ns):
        """Extract and concatenate all text, while respecting different tags like <a:br>."""
        full_text = []
        for paragraph in root.findall('.//a:p', ns):
            paragraph_text = []
            for element in paragraph:
                if element.tag == f"{{{ns['a']}}}r":
                    paragraph_text.append(self.format_drawing_text(element, ns))
                elif element.tag == f"{{{ns['a']}}}br":
                    paragraph_text.append('\n')

            # Join the paragraph content and add it to the full text
            full_text.append(''.join(paragraph_text))

        # Return the entire text block joined by new lines
        return '\n'.join(full_text).strip()

    def extract_drawing_metadata(self, zip_ref, relationships):
        """Extract drawing metadata based on the relationship IDs."""
        drawings = []
        for r_id, target in relationships.items():
            if target.startswith('../drawings/'):
                drawing_file = target.replace('../', 'xl/')
                try:
                    with zip_ref.open(drawing_file) as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        ns = {
                            'a': self.ns['drawings_main'],
                            'xdr': self.ns['drawings_sd']
                        }

                        # Extract all text as a continuous block while respecting line breaks
                        text_content = self.extract_full_text(root, ns)

                        if text_content:
                            drawings.append(f"Drawing (id: {r_id}):\n{text_content}")
                        else:
                            drawings.append(f"Drawing (id: {r_id}) from {drawing_file} (media)")
                except KeyError:
                    pass
        return drawings

    @staticmethod
    def format_drawing_text(run, ns):
        """Format the text in drawings, handling superscript and subscript, preserving spaces."""
        t_element = run.find('a:t', ns)
        if t_element is not None and t_element.text is not None:
            text = t_element.text
            r_pr = run.find('a:rPr', ns)
            baseline = int(r_pr.attrib.get('baseline', 0)) if r_pr is not None else 0

            # Apply formatting only to non-space characters
            if baseline > 0:  # Superscript
                formatted_text = ''.join(f"^{char}^" if char != ' ' else char for char in text)
            elif baseline < 0:  # Subscript
                formatted_text = ''.join(f"~{char}~" if char != ' ' else char for char in text)
            else:
                formatted_text = text

            return formatted_text
        return ""

    def convert(self):
        """Main conversion function to handle the Excel to Markdown transformation."""
        with zipfile.ZipFile(self.excel_file, 'r') as zip_ref:
            self.extract_shared_strings(zip_ref)
            self.extract_styles(zip_ref)
            worksheet_path = 'xl/worksheets/sheet1.xml'
            with zip_ref.open(worksheet_path) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                return self.build_markdown_table(
                    root,
                    self.extract_merged_cells(root),
                    self.extract_drawing_metadata(
                       zip_ref,
                       self.extract_relationships(zip_ref, worksheet_path)
                    )
                )

    def build_markdown_table(self, sheet_root, merged_cells, drawing_metadata):
        """Build the Markdown table from the sheet's XML structure."""
        markdown_table = []
        if drawing_metadata:
            markdown_table.append("\nDrawings:")
            markdown_table.extend(f"* {drawing}" for drawing in drawing_metadata)
            markdown_table.append("\n")


        for row in sheet_root.findall('main:sheetData/main:row', self.ns):
            row_data = [
                self.process_cell(cell, merged_cells)
                for cell in row.findall('main:c', self.ns)
            ]
            if row_data:
                markdown_table.append(f"| {' | '.join(row_data)} |")

        return self.add_table_header(markdown_table)

    def process_cell(self, cell, merged_cells):
        """Process each cell for Markdown conversion."""
        value = ""
        if cell.attrib.get('t') == 's':
            value = self.shared_strings[int(cell.find('main:v', self.ns).text)]
        elif cell.find('main:v', self.ns) is not None:
            value = cell.find('main:v', self.ns).text

        # Handle inline strings (is) that may contain rich text
        rich_text_elements = cell.findall('main:is/main:r', self.ns)
        if rich_text_elements:
            value = ''.join(
                t.text or '' for r in rich_text_elements for t in r.findall('main:t', self.ns)
            )

        # Apply styles such as superscript, subscript, and indentation
        style_idx = int(cell.attrib.get('s', -1))
        style = self.styles.get(style_idx, {})
        value = self.apply_styles(value, style)

        # Mark as merged if applicable
        cell_ref = cell.attrib['r']
        if any(cell_ref in m for m in merged_cells):
            value += " (merged)"
        return value or ""

    @staticmethod
    def apply_styles(value, style):
        """Apply styles such as superscript, subscript, and indentation."""
        if 'vertAlign' in style:
            if style['vertAlign'] == 'superscript':
                value = f"{value}"
            elif style['vertAlign'] == 'subscript':
                value = f"_{value}_"
        if 'indent' in style:
            value = f"{'-' * style['indent']}> {value}"
        return value

    @staticmethod
    def add_table_header(markdown_table):
        """Add a header to the Markdown table."""
        if markdown_table:
            header_separator = "| " + " | ".join([
                "-" * len(col) for col in markdown_table[0].split('|')[1:-1]
                ]) + " |"
            markdown_table.insert(1, header_separator)
        return "\n".join(markdown_table)


def main():
    """Main function to run the Excel to Markdown conversion."""
    if len(sys.argv) < 2:
        print("Usage: python excel2markdown.py <excel_file.xlsx>")
        sys.exit(1)

    excel_file = sys.argv[1]
    if not os.path.exists(excel_file):
        print(f"Error: File '{excel_file}' not found.")
        sys.exit(1)

    converter = ExcelToMarkdownConverter(excel_file)
    print(converter.convert())


if __name__ == "__main__":
    main()
  
