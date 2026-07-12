# item_parser.py
import xml.etree.ElementTree as ET
from fg_parser import clean_xml_text

def get_item_catalog(root):
    """
    Harvests all items from the master <item> node dynamically.
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
                continue

            item_type = clean_xml_text(entry.find("item_type"))
            
            # Line logging removed to minimize cumbersome large file dumps
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
                "damage": clean_xml_text(entry.find("damage")) or "0",
                "damage_dr": clean_xml_text(entry.find("damage_dr")) or "0",
                "range": clean_xml_text(entry.find("range")) or "N/A",
                "payload": clean_xml_text(entry.find("payload")) or "None",
                "training": clean_xml_text(entry.find("training")) or "None",
                "protection": clean_xml_text(entry.find("protection")),
                "protectionType": clean_xml_text(entry.find("protectionType")),
                "passive": clean_xml_text(entry.find("passive")),
                "active": clean_xml_text(entry.find("active")),
                "usage": clean_xml_text(entry.find("usage")),
                "raw_node": entry
            }
            
            if desc_node is not None:
                p_tag = desc_node.find("p")
                item_data["item_description"] = clean_xml_text(p_tag) if p_tag is not None else ""
            else:
                item_data["item_description"] = ""
                
            items_list.append(item_data)

        except Exception as item_err:
            print(f"  [!] CRITICAL ERROR parsing Item Node {idx}: {item_err}")
            continue

    print(f"[DEBUG END: ITEM PARSER] Successfully parsed and packaged {len(items_list)} items.")
    return sorted(items_list, key=lambda x: x["name"].lower())