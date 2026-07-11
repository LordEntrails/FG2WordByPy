import os
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
import fg_parser
import word_renderer
from ruleset_factory import RulesetFactory

# Master File and Global Output Targets
MOD_FILE = "FS_MtN.mod"
OUTPUT_FILE = "Mission_to_Nagol_Campaign.docx"

def main():
    try:
        if not os.path.exists(MOD_FILE):
            print(f"[!] Error: Could not find module package: '{MOD_FILE}'")
            return

        print(f"--- Initiating Fantasy Grounds Module Extraction Pipeline ---")
        print(f"Target Archive: {MOD_FILE}")

        # Open the module archive container
        mod_zip = zipfile.ZipFile(MOD_FILE, 'r')
        
        # =====================================================================
        # 1. VALIDATE SYSTEM METADATA & INITIALIZE FACTORY (definition.xml)
        # =====================================================================
        try:
            with mod_zip.open('definition.xml') as def_file:
                def_tree = ET.parse(def_file)
                def_root = def_tree.getroot()
                
            ruleset_node = def_root.find('ruleset')
            ruleset_name = ruleset_node.text.strip() if ruleset_node is not None else "Unknown"
        except KeyError:
            print("[WARNING] definition.xml missing from archive package. Defaulting validation pass.")
            ruleset_name = "Unknown"

        print(f"[SYSTEM] Module metadata verified. Active Ruleset: '{ruleset_name}'")
        
        # Instantiate the dynamic routing factory
        rf = RulesetFactory(ruleset_name)

        # =====================================================================
        # 2. LOAD CAMPAIGN DATABASE (db.xml)
        # =====================================================================
        try:
            xml_file = mod_zip.open('db.xml')
            db_root = ET.fromstring(xml_file.read())
            print("--- Successfully loaded the Fantasy Grounds XML data! ---")
        except KeyError:
            print("[!] Critical Error: db.xml was not found inside the module archive.")
            mod_zip.close()
            return

        # Open rule-specific Master Blueprint Template and wipe temporary text frames
        story_template = rf.get_template_path("story")  # Dynamically points to 'FS_Story_Template.docx'
        if not os.path.exists(story_template):
            print(f"[!] Critical Error: Core layout blueprint template missing: '{story_template}'")
            mod_zip.close()
            return

        doc = Document(story_template)
        for paragraph in doc.paragraphs:
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
            
        doc.add_heading("Campaign Narrative & Roster", level=1)
        
        # =====================================================================
        # 3. WRITE BOOK STRUCTURES SEQUENTIALLY
        # =====================================================================
        print(f"Processing structured narrative hierarchy...")
        story_pages = fg_parser.harvest_ordered_story_pages(db_root)
        
        for element in story_pages:
            item_type = element[0]
            
            if item_type == 'chapter':
                doc.add_heading(element[1], level=1)
                
            elif item_type == 'subchapter':
                doc.add_heading(element[1], level=2)
                
            elif item_type == 'page':
                page_name = element[1]
                page_node = element[2]
                
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
                                p = doc.add_paragraph(style=word_renderer.get_safe_style(doc, 'Heading 4'))
                                p.add_run(header_text)
                        
                        text_node = block.find("text[@type='formattedtext']")
                        if text_node is not None:
                            word_renderer.write_formatted_text(text_node, doc, block_type=block_type)
                            
                        if block_type == 'image' or block.find('.//image') is not None:
                            bitmap_node = block.find('.//bitmap')
                            caption_node = block.find("caption[@type='string']")
                            
                            if bitmap_node is not None and bitmap_node.text:
                                image_path = bitmap_node.text.strip()
                                success = word_renderer.extract_and_insert_image(mod_zip, image_path, doc)
                                
                                if success and caption_node is not None and caption_node.text:
                                    p_cap = doc.add_paragraph(style=word_renderer.get_safe_style(doc, 'Caption'))
                                    p_cap.alignment = 1
                                    p_cap.add_run(fg_parser.clean_xml_text(caption_node))
                                
        # =====================================================================
        # 4. WRITE MASTER NPC DOSSIERS
        # =====================================================================
        npcs = fg_parser.get_npc_catalog(db_root)
        if npcs:
            import npc_renderer
            print(f"\n[DEBUG] Total Categories Found: {len(npcs)}")
            for c in npcs:
                display_name = c['name'] if not c['is_unassigned'] else "Uncategorized/Loose"
                print(f"  -> Category: '{display_name}' (Contains {len(c['npcs'])} NPCs)")
                
            npc_template = rf.get_template_path("npc")  # Resolves dynamically to 'FS_NPC_Template.docx'
            npc_renderer.render_npc_appendix(mod_zip, npcs, doc, npc_template)

        # =====================================================================
        # 4.5 WRITE EQUIPMENT & ITEM LOGS
        # =====================================================================
        items_data = fg_parser.get_item_catalog(db_root)
        if items_data:
            import item_renderer
            item_template = rf.get_template_path("item")  # Resolves dynamically to 'FS_Item_Template.docx'
            item_renderer.render_item_appendix(mod_zip, items_data, doc, item_template)
        else:
            print("[INFO] No inventory items detected inside the campaign module.")
        
        # =====================================================================
        # 4.6 WRITE VEHICLE & HULL ENTRIES (APPENDIX C)
        # =====================================================================
        vehicles_data = rf.execute_parser("vehicle", db_root)  # Dynamically executes fs_vehicle_parser
        if vehicles_data:
            import vehicle_renderer
            vehicle_template = rf.get_template_path("vehicle")  # Resolves dynamically to 'FS_Vehicle_Template.docx'
            vehicle_renderer.render_vehicle_appendix(
                mod_zip=mod_zip, 
                vehicles_data=vehicles_data, 
                master_doc=doc, 
                template_path=vehicle_template
            )
        else:
            print("[INFO] No data objects populated. Vehicle pipeline bypassed.")

        # =====================================================================
        # 5. SAVE AND CLOSE
        # =====================================================================
        mod_zip.close()
        doc.save(OUTPUT_FILE)
        print("\n--- SUCCESS ---")
        print(f"Generated structured module workbook: {OUTPUT_FILE}")
        
    except Exception as e:
        import traceback
        print(f"\nProcessing Pipeline Failure: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()