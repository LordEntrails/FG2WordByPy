import os
import io
from docx.shared import Inches
from docxtpl import DocxTemplate, RichText
from PIL import Image as PILImage

def create_item_image_subdoc(tpl, mod_zip, asset_path, target_width=3.5):
    """
    Creates an inline sub-document for items using explicit dimension calculations
    to lock the true aspect ratio and prevent square distortion.
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
            
        subdoc = tpl.new_subdoc()
        p = subdoc.add_paragraph()
        p.alignment = 1  # Centered
        
        p.add_run().add_picture(output_stream, width=Inches(target_width), height=Inches(target_height))
        return subdoc

    except KeyError:
        return f"[MISSING MODULE ASSET: {asset_path}]"
    except Exception as e:
        print(f"  [!] ITEM ASSET WARNING: Failed to process picture '{asset_path}'. Error: {e}")
        return f"[CORRUPT ASSET: {asset_path}]"

# FIX: Added appendix_label="" as the 5th positional argument
def render_item_appendix(mod_zip, items_data, master_doc, template_path, appendix_label=""):
    """
    Compiles item lists using docxtpl and appends the structural table rows 
    into the master document stream.
    """
    if not os.path.exists(template_path):
        print(f"[!] Error: Item renderer blueprint missing: '{template_path}'")
        return False

    print(f"Compiling {len(items_data)} items into text layout tables...")

    tpl = DocxTemplate(template_path)

    for item in items_data:
        # 1. Handle the Image Asset Injection via Subdoc
        pic_path = item["raw_node"].findtext("picture") or ""
        item["picture"] = create_item_image_subdoc(tpl, mod_zip, pic_path, target_width=3.5)

        # 2. Handle RichText Descriptions
        desc_node = item["raw_node"].find("item_description")
        if desc_node is not None and len(desc_node.findall('p')) > 0:
            rt = RichText()
            paragraphs = desc_node.findall('p')
            for idx, p_node in enumerate(paragraphs):
                text = p_node.text.strip() if p_node.text else ""
                if text:
                    rt.add(text)
                    if idx < len(paragraphs) - 1:
                        rt.add("\n\n") 
            item["item_description"] = rt
        else:
            item["item_description"] = "No description cataloged."

    # FIX: Added appendix_label to the context payload
    context = {
        "items": items_data,
        "appendix_label": appendix_label
    }

    try:
        tpl.render(context)
        
        for element in tpl.element.body:
            master_doc.element.body.append(element)
            
        print("[SUCCESS] Equipment appendix compiled into chronicle workbook.")
        return True
        
    except Exception as render_err:
        print(f"[!] Item list single-pass layout parsing failed.")
        print(f"    Error Details: {render_err}")
        return False