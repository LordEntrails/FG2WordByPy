from fg_parser import clean_xml_text

def get_tables_catalog(root):
    """
    Extracts rollable mechanical random selection tables from the campaign tree,
    calculating dice expressions and packing row node configurations.
    """
    tables_list = []
    print("\n--- [DEBUG START: TABLES PARSER] ---")
    
    tables_root = root.find('tables')
    if tables_root is None:
        print("[!] TABLES DEBUG: No <tables> elements found inside this campaign root.")
        return []
        
    for idx, table_node in enumerate(tables_root.findall('*'), 1):
        t_name = clean_xml_text(table_node.find("name"))
        if not t_name:
            continue
            
        print(f"  [Table Node {idx}] Packaging configuration parameters for: '{t_name}'")
        
        # Calculate human-readable roll calculation metrics
        dice = clean_xml_text(table_node.find("dice")) or ""
        try:
            mod = int(clean_xml_text(table_node.find("mod")) or 0)
        except ValueError:
            mod = 0
            
        dice_str = dice if dice else "Default Scale"
        if mod > 0:
            dice_str += f" +{mod}"
        elif mod < 0:
            dice_str += f" {mod}"

        tables_list.append({
            "name": t_name,
            "dice_string": dice_str,
            "description": clean_xml_text(table_node.find("description")) or "No description cataloged.",
            "raw_node": table_node
        })
        
    print(f"[DEBUG END: TABLES PARSER] Wrapped {len(tables_list)} custom tables successfully.")
    return sorted(tables_list, key=lambda x: x["name"].lower())