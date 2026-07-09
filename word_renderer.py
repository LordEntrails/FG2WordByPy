import io
from docx import Document
from docx.shared import Inches
from PIL import Image
from fg_parser import clean_xml_text
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import RGBColor

def get_safe_style(doc, style_name, default_fallback='Normal'):
    if style_name in doc.styles:
        return style_name
    return default_fallback

def apply_inline_styles(node, paragraph_obj):
    """
    Recursively steps through text nodes to find inline <b> or <i> tags,
    carefully stripping code indents (\t) and raw line breaks (\n).
    """
    if node is None:
        return

    # Clean text before any child tags (removes raw code indents/newlines)
    if node.text:
        cleaned_text = node.text.lstrip('\n\t').rstrip('\n\t')
        if cleaned_text or node.text == " ":  # Keep intentional spaces
            paragraph_obj.add_run(node.text)
    
    for child in node:
        # Secure the text child data and strip any raw layout indentation
        child_text = child.text if child.text else ""
        cleaned_child_text = child_text.lstrip('\n\t').rstrip('\n\t')
        
        if child.tag == 'b':
            run = paragraph_obj.add_run(cleaned_child_text)
            run.bold = True
        elif child.tag == 'i':
            run = paragraph_obj.add_run(cleaned_child_text)
            run.italic = True
        else:
            run = paragraph_obj.add_run(cleaned_child_text)
            
        # Recurse if there are nested inline tags
        apply_inline_styles(child, paragraph_obj)
        
        # Clean text trailing immediately after a child tag closure
        if child.tail:
            cleaned_tail = child.tail.lstrip('\n\t').rstrip('\n\t')
            if cleaned_tail or child.tail == " ":
                paragraph_obj.add_run(child.tail)

def extract_and_insert_image(mod_zip, image_path, doc):
    try:
        with mod_zip.open(image_path) as img_file:
            image_stream = io.BytesIO(img_file.read())
        with Image.open(image_stream) as img:
            output_stream = io.BytesIO()
            img.save(output_stream, format="PNG")
            output_stream.seek(0)
            p_img = doc.add_paragraph()
            p_img.alignment = 1 # Centers the image frame alignment anchor
            # FIXED SIZE SCALE: Set to exactly 3.5 inches to fit two-columns
            p_img.add_run().add_picture(output_stream, width=Inches(3.5))
            return True
    except KeyError:
        p = doc.add_paragraph(style=get_safe_style(doc, 'Normal'))
        p.add_run(f"[MISSING IMAGE ASSET: {image_path}]").font.color.rgb = (200, 0, 0)
        return False

def write_formatted_text(xml_node, doc, block_type=None):
    if xml_node is None:
        return
    for child in xml_node:
        if child.tag == 'p':
            if block_type == 'frame' or xml_node.tag == 'frame':
                style = get_safe_style(doc, 'Chat Paragraph', 'Normal')
                p = doc.add_paragraph(style=style)
                # Clean frame paragraph text indents
                if child.text:
                    p.add_run(child.text.lstrip('\n\t'))
                apply_inline_styles(child, p)
            else:
                p = doc.add_paragraph(style=get_safe_style(doc, 'Normal'))
                # Clean standard paragraph text indents
                if child.text:
                    p.add_run(child.text.lstrip('\n\t'))
                apply_inline_styles(child, p)
                
        elif child.tag == 'h':
            p = doc.add_paragraph(style=get_safe_style(doc, 'Heading 5'))
            p.add_run(clean_xml_text(child))
            
        elif child.tag == 'frame':
            # Handle standalone <frame> text entries using your custom border style
            style = get_safe_style(doc, 'Chat Paragraph', 'Normal')
            p = doc.add_paragraph(style=style)
            apply_inline_styles(child, p)
            
        elif child.tag == 'li':
            p = doc.add_paragraph(style=get_safe_style(doc, 'List Paragraph'))
            p.add_run("•\t")
            apply_inline_styles(child, p)
            
        elif child.tag in ['linklist', 'link']:
            style = get_safe_style(doc, 'Internal Link', 'Normal')
            links = child.findall('link') if child.tag == 'linklist' else [child]
            for link in links:
                p = doc.add_paragraph(style=style)
                p.add_run("o\t")
                p.add_run(clean_xml_text(link))
                
        elif child.tag == 'table':
            xml_rows = child.findall('tr')
            if xml_rows:
                max_cols = max((len(r.findall('td')) + len(r.findall('th'))) for r in xml_rows)
                word_table = doc.add_table(rows=len(xml_rows), cols=max_cols)
                
                # 1. Apply your native Word table style
                word_table.style = 'Grid Table 4 Accent 1'
                
                # 2. Force Table Style Options Checkboxes:
                # Turn Header Row and Banded Rows ON, turn everything else OFF
                word_table.header_row = True
                word_table.row_banding = True
                word_table.first_column = False
                word_table.last_column = False
                word_table.column_banding = False
                
                for r_idx, xml_row in enumerate(xml_rows):
                    cells = xml_row.findall('td') + xml_row.findall('th')
                    for c_idx, cell in enumerate(cells):
                        if c_idx >= max_cols:
                            continue
                            
                        word_cell = word_table.cell(r_idx, c_idx)
                        word_cell.text = clean_xml_text(cell)
                        
                        for paragraph in word_cell.paragraphs:
                            paragraph.style = get_safe_style(doc, 'Normal')
                            
                            # Keep the header text bold manually since FG inputs 
                            # don't distinguish th from td visually.
                            if r_idx == 0:
                                for run in paragraph.runs:
                                    run.bold = True