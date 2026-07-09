import io
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from fg_parser import clean_xml_text
from word_renderer import write_formatted_text

def process_npc_graphics(tpl, mod_zip, npc):
    """
    Attaches images directly to the individual NPC data structure 
    so they can be called via {{npc.picture}} inside the loop.
    """
    raw_node = npc["raw_node"]
    pic_path = raw_node.findtext("picture") or ""
    token_flat_path = raw_node.findtext("token") or ""
    token_cam_path = raw_node.findtext("token3Dflat") or ""

    if pic_path:
        if "@" in pic_path:
            npc['picture'] = f"[IMAGE SOURCED EXTERNALLY: {pic_path}]"
        else:
            try:
                with mod_zip.open(pic_path) as img_file:
                    npc['picture'] = InlineImage(tpl, io.BytesIO(img_file.read()), width=Inches(3.5))
            except KeyError:
                npc['picture'] = f"[MISSING MODULE IMAGE: {pic_path}]"
    else:
        npc['picture'] = ""

    if token_flat_path:
        if "@" in token_flat_path:
            npc['token_flat'] = f"[TOKEN SOURCED EXTERNALLY: {token_flat_path}]"
        else:
            try:
                with mod_zip.open(token_flat_path) as img_file:
                    npc['token_flat'] = InlineImage(tpl, io.BytesIO(img_file.read()), width=Inches(2.0))
            except KeyError:
                npc['token_flat'] = f"[MISSING MODULE TOKEN: {token_flat_path}]"
    else:
        npc['token_flat'] = ""

    if token_cam_path:
        if "@" in token_cam_path:
            npc['token_camera'] = f"[3D TOKEN SOURCED EXTERNALLY: {token_cam_path}]"
        else:
            try:
                with mod_zip.open(token_cam_path) as img_file:
                    npc['token_camera'] = InlineImage(tpl, io.BytesIO(img_file.read()), width=Inches(2.0))
            except KeyError:
                npc['token_camera'] = f"[MISSING MODULE 3D TOKEN: {token_cam_path}]"
    else:
        npc['token_camera'] = ""

    return npc

def render_npc_appendix(mod_zip, structured_categories, doc_base, blueprint_path):
    """
    Renders multi-tiered categorized NPC lists into your master document.
    """
    if not structured_categories:
        return

    print(f"Injecting {len(structured_categories)} sorted category tracks into blueprints...")
    tpl = DocxTemplate(blueprint_path)

    # Run through groups to pre-populate image formats and text boxes elements
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

    # Simply hand the array list over to Jinja
    context = {
        "categories": structured_categories
    }

    try:
        tpl.render(context)
        doc_base.add_page_break()
        for element in tpl.element.body:
            doc_base.element.body.append(element)
    except Exception as jinja_err:
        print(f"\n[!] Category list single-pass layout parsing failed.")
        print(f"    Error Details: {jinja_err}")
        raise jinja_err