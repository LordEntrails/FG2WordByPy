from fg_parser import clean_xml_text

def get_starships_catalog(root):
    """
    Harvests starship profiles from the <starships> root node, structuring them 
    by category envelopes or unassigned asset branches.
    """
    structured_categories = []
    
    print("\n--- [DEBUG START: STARSHIPS PARSER] ---")
    
    starships_root = root.find('starships')
    if starships_root is None:
        print("[!] STARSHIPS DEBUG: No <starships> element blocks found inside this database.")
        return []

    def parse_starship_node(entry_node):
        ship_name = clean_xml_text(entry_node.find("name"))
        if not ship_name:
            return None
            
        # print(f"  [Starship Data] Processing Hull Profile: '{ship_name}'")
        
        return {
            "name": ship_name,
            "frame": clean_xml_text(entry_node.find("frame")) or "Standard Hull",
            "powercore": clean_xml_text(entry_node.find("powercore")) or "Parabattery Array",
            "crew": clean_xml_text(entry_node.find("crew")) or "1",
            "passengers": clean_xml_text(entry_node.find("passengers")) or "0",
            "pulse": clean_xml_text(entry_node.find("pulse")) or "0",
            "nova": clean_xml_text(entry_node.find("nova")) or "0",
            "handling": clean_xml_text(entry_node.find("handling")) or "0",
            "atmo": clean_xml_text(entry_node.find("atmo")) or "0",
            "computer": clean_xml_text(entry_node.find("computer")) or "0",
            "scanners": clean_xml_text(entry_node.find("scanners")) or "0",
            "targeting": clean_xml_text(entry_node.find("targeting")) or "0",
            "hull_points": clean_xml_text(entry_node.find("hull_points")) or "0",
            "structure": clean_xml_text(entry_node.find("structure")) or "0",
            "subsystem": clean_xml_text(entry_node.find("subsystem")) or "0",
            "capacitors": clean_xml_text(entry_node.find("capacitors")) or "0",
            "recharge": clean_xml_text(entry_node.find("recharge")) or "0",
            "cargo": clean_xml_text(entry_node.find("cargo")) or "0",
            "dockingspace": clean_xml_text(entry_node.find("dockingspace")) or "0",
            "fuel": clean_xml_text(entry_node.find("fuel")) or "0",
            "provisions": clean_xml_text(entry_node.find("provisions")) or "0",
            "scoops": clean_xml_text(entry_node.find("scoops")) or "0",
            "hardpoints_max": clean_xml_text(entry_node.find("hardpoints_max")) or "0",
            "integration": clean_xml_text(entry_node.find("integration")) or "0",
            "integration_max": clean_xml_text(entry_node.find("integration_max")) or "0",
            "raw_node": entry_node
        }

    # Gather uncategorized base records
    unassigned_bag = []
    for child in starships_root.findall('*'):
        if child.tag != 'category':
            ship_payload = parse_starship_node(child)
            if ship_payload:
                unassigned_bag.append(ship_payload)

    if unassigned_bag:
        structured_categories.append({
            "name": "Uncategorized/Loose",
            "is_unassigned": True,
            "starships": sorted(unassigned_bag, key=lambda x: x["name"].lower())
        })

    # Gather grouped subcategories
    for category_node in starships_root.findall('category'):
        cat_name = category_node.get('name') or "Auxiliary Class"
        cat_bag = []
        
        for child in category_node.findall('*'):
            ship_payload = parse_starship_node(child)
            if ship_payload:
                cat_bag.append(ship_payload)
                
        if cat_bag:
            structured_categories.append({
                "name": cat_name,
                "is_unassigned": False,
                "starships": sorted(cat_bag, key=lambda x: x["name"].lower())
            })

    print(f"[DEBUG END: STARSHIPS PARSER] Successfully compiled {len(structured_categories)} starship categories.")
    return structured_categories