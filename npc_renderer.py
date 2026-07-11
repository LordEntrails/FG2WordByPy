import io
import os
from docx.shared import Inches
from docxtpl import DocxTemplate
from PIL import Image as PILImage
from word_renderer import write_formatted_text

def create_image_subdoc(tpl, mod_zip, asset_path, target_width):
    if not asset_path or "@" in asset_path:
        return ""
    try:
        with mod_zip.open(asset_path) as img_file:
            img_data = img_file.read()
        with PILImage.open(io.BytesIO(img_data)) as img:
            width_px, height_px = img.size
            aspect_ratio = height_px / width_px
            target_height = target_width * aspect_ratio
            output_stream = io.BytesIO()
            img.save(output_stream, format="PNG")
            output_stream.seek(0)
            
        subdoc = tpl.new_subdoc()
        p = subdoc.add_paragraph()
        p.alignment = 1
        p.add_run().add_picture(output_stream, width=Inches(target_width), height=Inches(target_height))
        return subdoc
    except Exception:
        return ""

def render_npc_appendix(mod_zip, structured_categories, doc_base, blueprint_path, appendix_label=""):
    if not structured_categories or not os.path.exists(blueprint_path):
        return

    print(f"Injecting {len(structured_categories)} sorted category tracks into blueprints...")
    tpl = DocxTemplate(blueprint_path)

    for group in structured_categories:
        for npc in group["npcs"]:
            raw_node = npc["raw_node"]
            npc['picture'] = create_image_subdoc(tpl, mod_zip, raw_node.findtext("picture") or "", 3.5)
            npc['token_flat'] = create_image_subdoc(tpl, mod_zip, raw_node.findtext("token") or "", 1.7)
            npc['token_camera'] = create_image_subdoc(tpl, mod_zip, raw_node.findtext("token3Dflat") or "", 1.7)
            
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