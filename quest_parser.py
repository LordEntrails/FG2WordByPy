from fg_parser import clean_xml_text

def get_quests_catalog(root):
    """
    Harvests all generic quest entries from the campaign tree,
    extracting titles and storing data nodes for subdoc processing.
    """
    quests_list = []
    print("\n--- [DEBUG START: QUESTS PARSER] ---")
    
    # Locate the core quest wrapper root entry safely
    quests_root = root.find('quest') or root.find('.//quest')
    if quests_root is None:
        print("[!] QUEST DEBUG: No <quest> elements cataloged in this XML database.")
        return []
        
    all_nodes = quests_root.findall('*')
    print(f"[DEBUG] Detected {len(all_nodes)} raw records inside the <quest> tree.")

    for idx, entry in enumerate(all_nodes, 1):
        q_name = clean_xml_text(entry.find("name"))
        if not q_name:
            continue
            
        # print(f"  [Quest Node {idx}] Packaging tracking logs for: '{q_name}'")
        
        quests_list.append({
            "name": q_name,
            "raw_node": entry
        })
        
    print(f"[DEBUG END: QUESTS PARSER] Finished processing. Total packaged: {len(quests_list)} quests.")
    return sorted(quests_list, key=lambda x: x["name"].lower())