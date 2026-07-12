import os
from docxtpl import DocxTemplate
from word_renderer import write_formatted_text

def render_quests_appendix(mod_zip, quests_data, master_doc, template_path, appendix_label=""):
    if not os.path.exists(template_path):
        print(f"[!] Error: Quest blueprint layout missing: '{template_path}'")
        return False
        
    print(f"Compiling {len(quests_data)} mission dossiers into appendix directories...")
    tpl = DocxTemplate(template_path)

    for quest in quests_data:
        raw = quest["raw_node"]
        
        # 1. Format Player Facing Briefings using the native formatting engine
        text_node = raw.find("description")
        if text_node is not None:
            subdoc_text = tpl.new_subdoc()
            # Passing the subdocument as the 'doc' target routes tables natively to it
            write_formatted_text(text_node, subdoc_text, xml_root=raw)
            quest["quest_text"] = subdoc_text
        else:
            quest["quest_text"] = "No mission profile text logs archived."

        # 2. Format GM Classified Files using the native formatting engine
        gm_node = raw.find("gmnotes")
        if gm_node is not None and (gm_node.text or list(gm_node)):
            subdoc_gm = tpl.new_subdoc()
            write_formatted_text(gm_node, subdoc_gm, xml_root=raw)
            quest["gm_notes"] = subdoc_gm
        else:
            quest["gm_notes"] = None

    context = {
        "quests": quests_data,
        "appendix_label": appendix_label
    }

    try:
        tpl.render(context)
        master_doc.add_page_break()
        
        # Filter out section property blocks to prevent XML corruption
        for element in tpl.element.body:
            if hasattr(element, 'tag') and element.tag.endswith('sectPr'):
                continue
            master_doc.element.body.append(element)
            
        print("[SUCCESS] Quest manifest reference index compiled successfully with styled tables.")
        return True
    except Exception as render_err:
        print(f"[!] Quest appendix generation failed. Error Details: {render_err}")
        return False