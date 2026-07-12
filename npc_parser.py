# npc_parser.py
import xml.etree.ElementTree as ET
from collections import defaultdict
from fg_parser import clean_xml_text

def get_npc_catalog(root):
    """
    Gathers all valid NPC nodes by explicitly targeting top-level entries and category folders[cite: 24].
    """
    category_map = defaultdict(list)
    npc_directory = root.find('npc') or root.find('.//npc')
    
    if npc_directory is None:
        print("\n[!] CRITICAL ERROR: The <npc> database node was not found anywhere in db.xml.")
        return []

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

    for child in npc_directory:
        if child.tag == 'category':
            cat_name = child.get('name', '').strip()
            for sub_child in child:
                if sub_child.tag != 'categoryname':
                    process_npc_element(sub_child, cat_name)
        else:
            process_npc_element(child, "")

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