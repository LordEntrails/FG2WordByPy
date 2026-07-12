from fg_parser import clean_xml_text

def get_location_catalog(root):
    """
    Harvests all system, planetary, and site location nodes from the campaign tree,
    retaining narrative blocks and GM coordinates with verbose console tracking.
    """
    locations_list = []
    
    print("\n--- [DEBUG START: LOCATION PARSER] ---")
    
    # Locate the root location directory wrapper safely
    location_directory = root.find('location') or root.find('.//location')
    if location_directory is None:
        print("[!] LOCATION DEBUG: No <location> elements cataloged in this XML root.")
        return []

    all_nodes = location_directory.findall('*')
    print(f"[DEBUG] Detected {len(all_nodes)} raw records inside <location> tree.")

    for idx, entry in enumerate(all_nodes, 1):
        loc_name = clean_xml_text(entry.find("name"))
        if not loc_name:
            continue

        # print(f"  [Location Node {idx}] Processing Sector: '{loc_name}'")

        location_data = {
            "name": loc_name,
            "type": clean_xml_text(entry.find("type")) or "Unknown Structure",
            "raw_node": entry
        }
        locations_list.append(location_data)

    print(f"[DEBUG END: LOCATION PARSER] Finished processing. Total packaged: {len(locations_list)} locations.")
    return sorted(locations_list, key=lambda x: x["name"].lower())