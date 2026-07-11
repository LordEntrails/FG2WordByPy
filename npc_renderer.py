import io
import docx
from docx.shared import Inches
from docxtpl import DocxTemplate
from PIL import Image as PILImage
from fg_parser import clean_xml_text
from word_renderer import write_formatted_text

def create_image_subdoc(tpl, mod_zip, asset_path, target_width):
    """
    Creates an inline sub-document using explicit dimension calculations
    to overcome the sub-document environment constraint, locking true aspect ratios.
    """
    if not asset_path:
        return ""
        
    if "@" in asset_path:
        return f"[ASSET SOURCED EXTERNALLY: {asset_path}]"

    try:
        with mod_zip.open(asset_path) as img_file:
            img_data = img_file.read()
            
        with PILImage.open(io.BytesIO(img_data)) as img:
            width_px, height_px = img.size
            
            # Calculate explicit aspect ratio parameters
            aspect_ratio = height_px / width_px
            target_height = target_width * aspect_ratio
            
            output_stream = io.BytesIO()
            img.save(output_stream, format="PNG")
            output_stream.seek(0)
            
        # Initialize a clean sub-document layout stream
        subdoc = tpl.new_subdoc()
        p = subdoc.add_paragraph()
        p.alignment = 1  # Centered alignment anchor
        
        # FIX: Provide BOTH width and height explicitly to override subdoc square defaults
        p.add_run().add_picture(output_stream, width=Inches(target_width), height=Inches(target_height))
        return subdoc

    except KeyError:
        return f"[MISSING MODULE ASSET: {asset_path}]"
    except Exception as e:
        print(f"  [!] ASSET WARNING: Failed to process picture '{asset_path}'. Error: {e}")
        return f"[CORRUPT ASSET: {asset_path}]"

def process_npc_graphics(tpl, mod_zip, npc):
    """
    Converts image data channels into auto-scaling sub-documents 
    instead of using docxtpl's rigid InlineImage parameters.
    """
    raw_node = npc["raw_node"]
    pic_path = raw_node.findtext("picture") or ""
    token_flat_path = raw_node.findtext("token") or ""
    token_cam_path = raw_node.findtext("token3Dflat") or ""

    # Generate isolated layout anchors for each token slot with specific widths
    npc['picture'] = create_image_subdoc(tpl, mod_zip, pic_path, target_width=3.5)
    npc['token_flat'] = create_image_subdoc(tpl, mod_zip, token_flat_path, target_width=1.7)
    npc['token_camera'] = create_image_subdoc(tpl, mod_zip, token_cam_path, target_width=1.7)

    return npc

def render_npc_appendix(mod_zip, structured_categories, doc_base, blueprint_path, appendix_label=""):
    """
    Renders multi-tiered categorized NPC lists into your master document.
    """
    if not structured_categories:
        return

    print(f"Injecting {len(structured_categories)} sorted category tracks into blueprints...")
    tpl = DocxTemplate(blueprint_path)

    for group in structured_categories:
        for npc in group["npcs"]:
            npc = process_npc_graphics(tpl, mod_zip, npc)
            
            notes_node = npc["raw_node"].find('notes')
            if notes_node is not None:
                subdoc = tpl.new_subdoc()
                write_formatted_text(notes_node, subdoc)
                npc["npc_notes"] = subdoc
            else:
                npc["npc_notes"] = ""

    # FIXED: Added the missing comma at the end of the first line below
    context = {
        "categories": structured_categories,
        "appendix_label": appendix_label  
    }

    try:
        tpl.render(context)
        doc_base.add_page_break()
        for element in tpl.element.body:
            doc_base.element.body.append(element)
    except Exception as render_err:
        import traceback
        print(f"\n[!] CRITICAL TEMPLATE RENDER FAILURE inside npc_renderer.py!")
        print(f"--- TRACEBACK LOGS ---")
        traceback.print_exc()
        print(f"----------------------")
        raise render_err