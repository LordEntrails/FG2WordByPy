import zipfile
import xml.etree.ElementTree as ET

def clean_xml_text(element):
    return element.text.strip() if element is not None and element.text else ""

def get_xml_root(mod_filename):
    """Opens the .mod file and returns the XML root and open zip reference."""
    mod_zip = zipfile.ZipFile(mod_filename, 'r')
    xml_file = mod_zip.open('db.xml')
    return ET.fromstring(xml_file.read()), mod_zip

def harvest_ordered_story_pages(root):
    """
    Parses the refmanualindex by strictly sorting chapters, subchapters,
    and refpages by their internal <order> numerical tags.
    Returns a flat list of structural layout items:
    [('chapter', 'Name'), ('subchapter', 'Name'), ('page', 'Name', page_node), ...]
    """
    elements = []
    ref_index = root.find('.//reference/refmanualindex')
    ref_data = root.find('.//reference/refmanualdata')
    
    if ref_index is None or ref_data is None:
        return elements

    # 1. Process and sort Chapters
    chapters_node = ref_index.find('chapters')
    if chapters_node is None:
        return elements
        
    ordered_chapters = []
    for chapter in chapters_node:
        order_val = chapter.find("order[@type='number']")
        order_num = int(order_val.text) if order_val is not None and order_val.text else 999
        name = clean_xml_text(chapter.find('name'))
        ordered_chapters.append((order_num, name, chapter))
    ordered_chapters.sort(key=lambda x: x[0])

    # Walk sorted Chapters
    for ch_order, ch_name, chapter_node in ordered_chapters:
        if ch_name:
            elements.append(('chapter', ch_name))
        
        # 2. Process and sort Sub-chapters
        subchapters_node = chapter_node.find('subchapters')
        if subchapters_node is None:
            continue
            
        ordered_subchapters = []
        for subch in subchapters_node:
            order_val = subch.find("order[@type='number']")
            order_num = int(order_val.text) if order_val is not None and order_val.text else 999
            name = clean_xml_text(subch.find('name'))
            ordered_subchapters.append((order_num, name, subch))
        ordered_subchapters.sort(key=lambda x: x[0])
        
        # Walk sorted Sub-chapters
        for subch_order, subch_name, subch_node in ordered_subchapters:
            if subch_name:
                elements.append(('subchapter', subch_name))
            
            # 3. Process and sort Individual Pages
            refpages_node = subch_node.find('refpages')
            if refpages_node is None:
                continue
                
            ordered_refpages = []
            for refpage in refpages_node:
                order_val = refpage.find("order[@type='number']")
                order_num = int(order_val.text) if order_val is not None and order_val.text else 999
                name = clean_xml_text(refpage.find('name'))
                ordered_refpages.append((order_num, name, refpage))
            ordered_refpages.sort(key=lambda x: x[0])
            
            # Walk sorted Pages and grab underlying text block nodes
            for pg_order, pg_name, pg_node in ordered_refpages:
                record_node = pg_node.find('.//recordname')
                if record_node is not None and record_node.text:
                    page_id = record_node.text.split('.')[-1]
                    page_data_node = ref_data.find(page_id)
                    
                    if page_data_node is not None and pg_name:
                        # Skip master directory placeholders
                        if pg_name.lower() == "npcs":
                            continue
                        elements.append(('page', pg_name, page_data_node))
                        
    return elements

from collections import defaultdict

def get_npc_catalog(root):
    """
    Gathers all valid NPC nodes by explicitly targeting top-level entries and 
    category folders, bypassing sub-list noise completely. Natively compatible 
    with standard xml.etree.ElementTree.
    """
    category_map = defaultdict(list)
    
    npc_directory = root.find('npc') or root.find('.//npc')
    if npc_directory is None:
        print("\n[!] CRITICAL ERROR: The <npc> database node was not found anywhere in db.xml.")
        return []

    # 1. Helper function to clean and map an individual NPC element node safely
    def process_npc_element(entry, cat_name):
        npc_name = clean_xml_text(entry.find("name"))
        if not npc_name:
            return

        hp_node = entry.find("bodypoints") or entry.find("hp")
        bp_val = clean_xml_text(hp_node) or "0"

        npc_stats = {
            "name": npc_name,
            "species": clean_xml_text(entry.find("species")),
            "bp": bp_val,
            "str": clean_xml_text(entry.find("strength")) or "0",
            "agl": clean_xml_text(entry.find("agility")) or "0",
            "per": clean_xml_text(entry.find("perception")) or "0",
            "crd": clean_xml_text(entry.find("coordination")) or "0",
            "int": clean_xml_text(entry.find("intelligence")) or "0",
            "wil": clean_xml_text(entry.find("willpower")) or "0",
            "move": clean_xml_text(entry.find("move")) or "0",
            "init": clean_xml_text(entry.find("initiative")) or "0",
            "score": clean_xml_text(entry.find("score")) or "0",
            "behavior": clean_xml_text(entry.find("behavior")) or "Cunning"
        }

        # Sub-list safe loops mapping (Free of attribute predicates)
        for list_name, xml_tag, field_mappings in [
            ("attacklist", "attacklist", [("hit", ["atk", "hit", "hitpercent"]), ("dmg", ["damage", "damage_dr"]), ("type", ["damagetype"])]),
            ("defenselist", "defenselist", [("prot", ["protection", "armor"]), ("type", ["protectiontype", "type"])]),
            ("skillslist", "skillslist", [("score", ["score", "skills_score"])])
        ]:
            extracted_items = []
            list_node = entry.find(xml_tag)
            if list_node is not None:
                for item in list_node.findall('*'):
                    item_data = {"name": clean_xml_text(item.find("name"))}
                    for field, xml_choices in field_mappings:
                        val = "0"
                        for choice in xml_choices:
                            node = item.find(choice)
                            if node is not None:
                                val = clean_xml_text(node)
                                break
                        item_data[field] = val
                    extracted_items.append(item_data)
            npc_stats[list_name] = extracted_items

        notes_node = entry.find('notes')
        npc_stats["notes"] = clean_xml_text(notes_node) if notes_node is not None else ""
        
        type_node = entry.find("npc_typevalue") or entry.find("npc_type")
        npc_stats["type"] = type_node.text.strip().lower() if type_node is not None and type_node.text else "minor"
        npc_stats["raw_node"] = entry
        
        category_map[cat_name].append(npc_stats)

    # 2. Iterate through immediate entries under <npc>
    for child in npc_directory:
        if child.tag == 'category':
            # This is a folder! Process its inner characters
            cat_name = child.get('name', '').strip()
            for sub_child in child:
                if sub_child.tag != 'categoryname': # Skip structural metadata strings
                    process_npc_element(sub_child, cat_name)
        else:
            # Loose uncategorized NPC at the base root
            process_npc_element(child, "")

    # 3. Structural Sort Assembly Pipeline
    structured_categories = []
    
    if "" in category_map:
        unassigned_sorted = sorted(category_map[""], key=lambda x: x["name"].lower())
        structured_categories.append({
            "name": "",
            "is_unassigned": True,
            "npcs": unassigned_sorted
        })

    sorted_group_names = sorted([k for k in category_map.keys() if k != ""], key=lambda s: s.lower())
    for cat_name in sorted_group_names:
        sorted_npcs = sorted(category_map[cat_name], key=lambda x: x["name"].lower())
        structured_categories.append({
            "name": cat_name,
            "is_unassigned": False,
            "npcs": sorted_npcs
        })

    return structured_categories

def get_item_catalog(root):
    """
    Harvests all items from the master <item> node dynamically with aggressive 
    terminal debug checkpoints to isolate pipeline failures.
    """
    items_list = []
    
    print("\n--- [DEBUG START: ITEM PARSER] ---")
    item_directory = root.find('item') or root.find('.//item')
    
    if item_directory is None:
        print("[!] ITEM DEBUG: No <item> or <.//item> node found in the XML root.")
        return []

    all_items = item_directory.findall('*')
    print(f"[DEBUG] Found {len(all_items)} total raw child nodes inside <item> branch.")

    for idx, entry in enumerate(all_items, 1):
        try:
            item_name = clean_xml_text(entry.find("name"))
            if not item_name:
                print(f"  [Item Node {idx}] Skipped: Missing <name> field.")
                continue

            item_type = clean_xml_text(entry.find("item_type"))
            print(f"  [Item Node {idx}] Processing: '{item_name}' (Type: {item_type})")

            desc_node = entry.find("item_description")
            subtype_val = clean_xml_text(entry.find("subtype")) or clean_xml_text(entry.find("veh_subtype")) or "N/A"
            type_val = clean_xml_text(entry.find("item_typevalue")) or ""

            item_data = {
                "name": item_name,
                "type": item_type,
                "typevalue": type_val.lower(),
                "subtype": subtype_val,
                "cost": clean_xml_text(entry.find("cost")) or "0",
                "weight": clean_xml_text(entry.find("weight")) or "0",
                
                # Weapon fields
                "damage": clean_xml_text(entry.find("damage")) or "0",
                "damage_dr": clean_xml_text(entry.find("damage_dr")) or "0",
                "range": clean_xml_text(entry.find("range")) or "N/A",
                "payload": clean_xml_text(entry.find("payload")) or "None",
                "training": clean_xml_text(entry.find("training")) or "None",
                
                # Defense fields
                "protection": clean_xml_text(entry.find("protection")),
                "protectionType": clean_xml_text(entry.find("protectionType")),
                
                # Scanner fields
                "passive": clean_xml_text(entry.find("passive")),
                "active": clean_xml_text(entry.find("active")),
                "usage": clean_xml_text(entry.find("usage")),
                
                "raw_node": entry
            }
            
            # Safe extraction of base paragraph descriptions
            if desc_node is not None:
                p_tag = desc_node.find("p")
                item_data["item_description"] = clean_xml_text(p_tag) if p_tag is not None else ""
            else:
                item_data["item_description"] = ""
                
            items_list.append(item_data)

        except Exception as item_err:
            print(f"  [!] CRITICAL ERROR parsing Item Node {idx} ('{entry.tag}'): {item_err}")
            continue

    print(f"[DEBUG END: ITEM PARSER] Successfully parsed and packaged {len(items_list)} items.")
    
    # Alphabetize items naturally by name
    return sorted(items_list, key=lambda x: x["name"].lower())