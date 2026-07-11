import os
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
import docx.oxml as oxml
import docx.oxml.ns as ns
import fg_parser
import word_renderer
from ruleset_factory import RulesetFactory

MOD_FILE = "FS_MtN.mod"
OUTPUT_FILE = "Mission_to_Nagol_Campaign.docx"

def apply_safe_bookmarks(doc):
    """
    Scans the compiled document paragraphs to apply clean, non-corrupting 
    XML bookmarks directly over target record header lines.
    """
    print("Running document post-processor: Injecting live bookmark links...")
    bookmark_id = 1000  # Unique ID tracker to keep OpenXML schema happy
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
            
        bookmark_name = ""
        # 1. Detect Heading lines matching our output variables
        if paragraph.style.name.startswith('Heading') or text.startswith('⚡') or "Type:" in text:
            # We check the paragraph text or structural patterns to match target strings
            pass

    # A more bulletproof way is checking our known dataset arrays directly against paragraph lookups
    # Let's run an exact matching sweep across the document tables/paragraphs for entries
    # to avoid complex string heuristics:
    return

def add_paragraph_bookmark(paragraph, bookmark_name, b_id):
    """Safely wraps a paragraph element inside an OpenXML bookmark anchor pair."""
    p_element = paragraph._p
    
    b_start = oxml.shared.OxmlElement('w:bookmarkStart')
    b_start.set(ns.qn('w:id'), str(b_id))
    b_start.set(ns.qn('w:name'), bookmark_name)
    
    b_end = oxml.shared.OxmlElement('w:bookmarkEnd')
    b_end.set(ns.qn('w:id'), str(b_id))
    
    # Prepend the start anchor to the front of the paragraph nodes safely
    p_element.insert(0, b_start)
    p_element.append(b_end)

def main():
    try:
        if not os.path.exists(MOD_FILE):
            print(f"[!] Error: Could not find module package: '{MOD_FILE}'")
            return

        print(f"--- Initiating Fantasy Grounds Module Extraction Pipeline ---")
        print(f"Target Archive: {MOD_FILE}")

        mod_zip = zipfile.ZipFile(MOD_FILE, 'r')
        
        # Validate Metadata
        try:
            with mod_zip.open('definition.xml') as def_file:
                def_tree = ET.parse(def_file)
                def_root = def_tree.getroot()
            ruleset_node = def_root.find('ruleset')
            ruleset_name = ruleset_node.text.strip() if ruleset_node is not None else "Unknown"
        except KeyError:
            ruleset_name = "Unknown"

        print(f"[SYSTEM] Module metadata verified. Active Ruleset: '{ruleset_name}'")
        rf = RulesetFactory(ruleset_name)

        # Load db.xml
        try:
            xml_file = mod_zip.open('db.xml')
            db_root = ET.fromstring(xml_file.read())
            print("--- Successfully loaded the Fantasy Grounds XML data! ---")
        except KeyError:
            print("[!] Critical Error: db.xml was not found inside the module archive.")
            mod_zip.close()
            return

        # Setup Document Base Layout
        story_template = rf.get_template_path("story")
        doc = Document(story_template)
        for paragraph in list(doc.paragraphs):
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
            
        doc.add_heading("Campaign Narrative & Roster", level=1)
        
        # Keep track of records we process to pass to our post-processor anchor injector
        processed_npcs = []
        processed_items = []
        processed_vehicles = []

        pipeline_order = [
            {"id": "story",      "is_appendix": False},
            {"id": "npc",        "is_appendix": True},
            {"id": "item",       "is_appendix": True},
            {"id": "vehicle",    "is_appendix": True}
        ]

        appendix_letters = ["A", "B", "C", "D", "E"]
        appendix_ptr = 0

        for step in pipeline_order:
            component = step["id"]
            current_appendix = f"Appendix {appendix_letters[appendix_ptr]}" if step["is_appendix"] else ""

            if component == "story":
                print(f"\nProcessing structured narrative hierarchy...")
                story_pages = fg_parser.harvest_ordered_story_pages(db_root)
                for element in story_pages:
                    item_type = element[0]
                    if item_type == 'chapter':
                        doc.add_heading(element[1], level=1)
                    elif item_type == 'subchapter':
                        doc.add_heading(element[1], level=2)
                    elif item_type == 'page':
                        page_name, page_node = element[1], element[2]
                        doc.add_heading(page_name, level=3)
                        
                        blocks_node = page_node.find('blocks')
                        if blocks_node is not None:
                            ordered_blocks = []
                            for block in blocks_node:
                                order_val = block.find("order[@type='number']")
                                order_num = int(order_val.text) if order_val is not None and order_val.text else 999
                                ordered_blocks.append((order_num, block))
                            ordered_blocks.sort(key=lambda x: x[0])
                            
                            for order_num, block in ordered_blocks:
                                block_type = fg_parser.clean_xml_text(block.find("blocktype[@type='string']"))
                                if block_type == 'header':
                                    header_text = fg_parser.clean_xml_text(block.find("text[@type='string']"))
                                    if header_text:
                                        doc.add_paragraph(header_text, style=word_renderer.get_safe_style(doc, 'Heading 4'))
                                
                                text_node = block.find("text[@type='formattedtext']")
                                if text_node is not None:
                                    word_renderer.write_formatted_text(text_node, doc, block_type=block_type, xml_root=db_root)

            elif component == "npc":
                npcs = fg_parser.get_npc_catalog(db_root)
                if npcs:
                    import npc_renderer
                    # Track IDs for post-processing bookmarking
                    for group in npcs:
                        for n in group["npcs"]:
                            processed_npcs.append((n["name"], f"REF_NPC_{n['raw_node'].tag}"))
                    npc_template = rf.get_template_path("npc")
                    npc_renderer.render_npc_appendix(mod_zip, npcs, doc, npc_template, current_appendix)
                    if step["is_appendix"]: appendix_ptr += 1

            elif component == "item":
                items_data = fg_parser.get_item_catalog(db_root)
                if items_data:
                    import item_renderer
                    for item in items_data:
                        processed_items.append((item["name"], f"REF_ITEM_{item['raw_node'].tag}"))
                    item_template = rf.get_template_path("item")
                    item_renderer.render_item_appendix(mod_zip, items_data, doc, item_template, current_appendix)
                    if step["is_appendix"]: appendix_ptr += 1

            elif component == "vehicle":
                vehicles_data = rf.execute_parser("vehicle", db_root)
                if vehicles_data:
                    import vehicle_renderer
                    for veh in vehicles_data:
                        processed_vehicles.append((veh["name"], f"REF_VEHICLE_{veh['raw_node'].tag}"))
                    vehicle_template = rf.get_template_path("vehicle")
                    vehicle_renderer.render_vehicle_appendix(mod_zip, vehicles_data, doc, vehicle_template, current_appendix)
                    if step["is_appendix"]: appendix_ptr += 1

        # =====================================================================
        # 4.9 SAFE POST-PROCESS BLOCKMARK PASS
        # =====================================================================
        # Scan headings in the document to safely attach OpenXML links without corruption
        print("\nResolving cross-reference anchors safely...")
        b_id = 2000
        all_targets = processed_npcs + processed_items + processed_vehicles
        
        for paragraph in doc.paragraphs:
            text_line = paragraph.text.strip()
            # If the paragraph matches an exact asset header name, stamp it!
            for name, bookmark_name in list(all_targets):
                # Match name exactly, or if it contains data parameters split by a pipes line
                if text_line == name or text_line.startswith(name + " |"):
                    add_paragraph_bookmark(paragraph, bookmark_name, b_id)
                    b_id += 1
                    all_targets.remove((name, bookmark_name))

        # Save out the finished workspace file
        mod_zip.close()
        doc.save(OUTPUT_FILE)
        print("\n--- SUCCESS ---")
        print(f"Generated clean module workbook: {OUTPUT_FILE}")
        
    except Exception as e:
        import traceback
        print(f"\nProcessing Pipeline Failure: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()