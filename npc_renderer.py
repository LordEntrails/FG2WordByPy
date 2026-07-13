# npc_renderer.py
import io
import os
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from PIL import Image as PILImage
from word_renderer import write_formatted_text

def create_image_subdoc(tpl, mod_zip, asset_path, target_width):
    if not asset_path:
        return ""
        
    raw_path = asset_path.split("@")[0].strip()
    clean_path = raw_path.replace("\\", "/")
    
    zip_contents = {name.lower(): name for name in mod_zip.namelist()}
    if clean_path.lower() not in zip_contents:
        return ""
        
    true_zip_path = zip_contents[clean_path.lower()]

    try:
        with mod_zip.open(true_zip_path) as img_file:
            img_data = img_file.read()
        with PILImage.open(io.BytesIO(img_data)) as img:
            output_stream = io.BytesIO()
            img.save(output_stream, format="PNG")
            output_stream.seek(0)
            
        # FIXED: Return a native template InlineImage token reference bound directly to the master template file!
        return InlineImage(tpl, output_stream, width=Inches(target_width))
    except Exception as err:
        print(f"  [!] PIL CONVERSION CRASH on file '{true_zip_path}': {err}")
        return ""

def render_npc_appendix(mod_zip, structured_categories, doc_base, blueprint_path, appendix_label=""):
    if not structured_categories or not os.path.exists(blueprint_path):
        return

    tpl = DocxTemplate(blueprint_path)

    for group in structured_categories:
        # Diagnostic print for categorization validation tracking
        print(f"\n[NPC TELEMETRY] Packaging Category Track: '{group['name'] or 'Uncategorized'}'")
        for npc in group["npcs"]:
            print(f" -> Extracting Character Sheet: '{npc['name']}'")
            raw_node = npc["raw_node"]
            
            # Look inside standard token/picture targets
            npc['picture'] = create_image_subdoc(tpl, mod_zip, raw_node.findtext("picture") or "", 3.5)
            npc['token_flat'] = create_image_subdoc(tpl, mod_zip, raw_node.findtext("token") or "", 1.2)
            npc['token_camera'] = create_image_subdoc(tpl, mod_zip, raw_node.findtext("token3Dflat") or "", 1.2)
            
            notes_node = raw_node.find('notes')
            if notes_node is not None:
                subdoc = tpl.new_subdoc()
                write_formatted_text(notes_node, subdoc)
                npc["npc_notes"] = subdoc
            else:
                npc["npc_notes"] = ""

    context = {"categories": structured_categories, "appendix_label": appendix_label}
    tpl.render(context)
    doc_base.add_page_break()
    for element in tpl.element.body:
        doc_base.element.body.append(element)