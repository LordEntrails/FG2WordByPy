import io
import docx.oxml as oxml
import docx.oxml.ns as ns
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

def insert_internal_hyperlink(paragraph, text, bookmark_name):
    """
    Inserts a native Microsoft Word cross-reference hyperlink to an internal bookmark.
    """
    hyperlink = oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(ns.qn('w:anchor'), bookmark_name)
    
    new_run = oxml.shared.OxmlElement('w:r')
    rPr = oxml.shared.OxmlElement('w:rPr')
    
    # Standard blue underline hyperlink text formatting look
    color = oxml.shared.OxmlElement('w:color')
    color.set(ns.qn('w:val'), '0000FF')
    rPr.append(color)
    
    u = oxml.shared.OxmlElement('w:u')
    u.set(ns.qn('w:val'), 'single')
    rPr.append(u)
    
    new_run.append(rPr)
    text_node = oxml.shared.OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    
    paragraph._p.append(hyperlink)
    return hyperlink

def embed_encounter_table(doc, record_name, xml_root):
    """
    Finds a <battle> record inside db.xml, extracts its combatant list,
    and writes a concise encounter table directly inline with live links.
    """
    if xml_root is None or not record_name:
        return False
        
    # Example record_name: "battle.id-00003" -> locate target node id-00003
    node_id = record_name.split('.')[-1] if '.' in record_name else record_name
    battle_node = xml_root.find(f"./battle/{node_id}")
    
    if battle_node is None:
        return False

    battle_name = clean_xml_text(battle_node.find("name")) or "Unnamed Skirmish"
    
    # Styled block header for our embedded encounter
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Inches(0.15)
    p_title.paragraph_format.space_after = Inches(0.05)
    p_title.add_run(f"⚔️ Battle Deployment: {battle_name}").bold = True
    
    npclist = battle_node.find("npclist")
    if npclist is None or not list(npclist):
        p = doc.add_paragraph()
        p.add_run("   (No combatants cataloged in this encounter structure)").italic = True
        return True

    # Construct the concise table structure
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Grid Table 4 Accent 1'
    table.header_row = True
    table.row_banding = True
    table.first_column = False
    table.last_column = False
    table.column_banding = False
    
    # Headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Qty"
    hdr_cells[1].text = "Combatant Designation"
    for cell in hdr_cells:
        for p in cell.paragraphs:
            p.style = get_safe_style(doc, 'Normal')
            for run in p.runs:
                run.bold = True

    # Pull each combatant out of the manifest list
    for child in npclist.findall('*'):
        count = clean_xml_text(child.find("count")) or "1"
        npc_name = clean_xml_text(child.find("name")) or "Unknown Combatant"
        
        # Pull the absolute reference to establish our internal appendix anchor
        npc_link_node = child.find("link")
        bookmark_anchor = ""
        if npc_link_node is not None:
            rec_ref = npc_link_node.get("recordname") or npc_link_node.text or ""
            if "npc.id-" in rec_ref:
                target_id = rec_ref.split('.')[-1]
                bookmark_anchor = f"REF_NPC_{target_id}"

        row_cells = table.add_row().cells
        row_cells[0].text = f"{count}x"
        
        p_name = row_cells[1].paragraphs[0]
        p_name.style = get_safe_style(doc, 'Normal')
        if bookmark_anchor:
            # Inject live cross-reference hyperlink point straight to the Appendix
            insert_internal_hyperlink(p_name, npc_name, bookmark_anchor)
        else:
            p_name.add_run(npc_name)
            
    # Add spacing paragraph below the table to clear formatting
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Inches(0.1)
    return True

def write_formatted_text(xml_node, doc, block_type=None, xml_root=None):
    if xml_node is None:
        return
    for child in xml_node:
        if child.tag == 'p':
            if block_type == 'frame' or xml_node.tag == 'frame':
                style = get_safe_style(doc, 'Chat Paragraph', 'Normal')
                p = doc.add_paragraph(style=style)
                if child.text:
                    p.add_run(child.text.lstrip('\n\t'))
                apply_inline_styles(child, p)
            else:
                p = doc.add_paragraph(style=get_safe_style(doc, 'Normal'))
                if child.text:
                    p.add_run(child.text.lstrip('\n\t'))
                apply_inline_styles(child, p)
                
        elif child.tag == 'h':
            p = doc.add_paragraph(style=get_safe_style(doc, 'Heading 5'))
            p.add_run(clean_xml_text(child))
            
        elif child.tag == 'frame':
            style = get_safe_style(doc, 'Chat Paragraph', 'Normal')
            p = doc.add_paragraph(style=style)
            apply_inline_styles(child, p)
            
        elif child.tag == 'li':
            p = doc.add_paragraph(style=get_safe_style(doc, 'List Paragraph'))
            p.add_run("•\t")
            apply_inline_styles(child, p)
            
        elif child.tag in ['linklist', 'link']:
            links = child.findall('link') if child.tag == 'linklist' else [child]
            for link in links:
                link_class = link.get("class") or ""
                record_name = link.get("recordname") or link.text or ""
                link_text = clean_xml_text(link)
                
                # INTERCEPT: If it's an encounter link, extract data and draw the inline roster table
                if link_class == "battle":
                    success = embed_encounter_table(doc, record_name, xml_root)
                    if success:
                        continue # Skip drawing standard bullet links for battles
                
                # FALLBACK/OTHER LINKS: Build future-proof dynamic appendix hyperlinks
                style = get_safe_style(doc, 'Internal Link', 'Normal')
                p = doc.add_paragraph(style=style)
                p.add_run("o\t")
                
                bookmark_name = ""
                if "npc.id-" in record_name:
                    bookmark_name = f"REF_NPC_{record_name.split('.')[-1]}"
                elif "item.id-" in record_name:
                    bookmark_name = f"REF_ITEM_{record_name.split('.')[-1]}"
                elif "vehicle.id-" in record_name:
                    bookmark_name = f"REF_VEHICLE_{record_name.split('.')[-1]}"
                
                if bookmark_name:
                    insert_internal_hyperlink(p, link_text, bookmark_name)
                else:
                    p.add_run(link_text)
                
        elif child.tag == 'table':
            xml_rows = child.findall('tr')
            if xml_rows:
                max_cols = max((len(r.findall('td')) + len(r.findall('th'))) for r in xml_rows)
                word_table = doc.add_table(rows=len(xml_rows), cols=max_cols)
                
                word_table.style = 'Grid Table 4 Accent 1'
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
                            if r_idx == 0:
                                for run in paragraph.runs:
                                    run.bold = True