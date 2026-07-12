from fg_parser import clean_xml_text

def get_vehicle_catalog(root):
    """
    Harvests all vehicle subnodes from the master xml tree, mapping every stat 
    including integrated components and narrative blocks cleanly with loud terminal diagnostics.[cite: 13]
    """
    vehicles_list = []
    
    print("\n--- [DEBUG START: VEHICLE PARSER] ---")
    
    # Locate the root vehicle block wrapper safely[cite: 13]
    vehicle_directory = root.find('vehicle') or root.find('.//vehicle')
    if vehicle_directory is None:
        print("[!] VEHICLE DEBUG: No <vehicle> or <.//vehicle> tags found inside this XML root!")
        return []

    all_nodes = vehicle_directory.findall('*')
    print(f"[DEBUG] Detected {len(all_nodes)} raw child records inside <vehicle> root element.")

    for idx, entry in enumerate(all_nodes, 1):
        veh_name = clean_xml_text(entry.find("name"))
        if not veh_name:
            print(f"  [Vehicle Node {idx}] Skipped: Child node is missing a structural <name> tag.")
            continue

        # print(f"  [Vehicle Node {idx}] Processing Target: '{veh_name}'")

        # Extract nested inventory components (like integrated GPS/Radios)[cite: 13]
        components_array = []
        inv_list_node = entry.find("inventorylist")
        if inv_list_node is not None:
            for item_node in inv_list_node.findall('*'):
                comp_name = clean_xml_text(item_node.find("name"))
                if comp_name:
                    components_array.append({
                        "name": comp_name,
                        "count": clean_xml_text(item_node.find("count")) or "1",
                        "location": clean_xml_text(item_node.find("location")) or "Integrated"
                    })

        # Remap the data payload to fit the actual XML elements[cite: 13]
        vehicle_data = {
            "name": veh_name,
            "type": clean_xml_text(entry.find("type")) or "Vehicle",
            "framesize": clean_xml_text(entry.find("framesize")) or "N/A",
            "cost": clean_xml_text(entry.find("cost")) or "0",
            "powercore": clean_xml_text(entry.find("powercore")) or "N/A",
            
            # System Capacities[cite: 13]
            "crew": clean_xml_text(entry.find("crew")) or "1",
            "passengers": clean_xml_text(entry.find("passengers")) or "0",
            "cargo": clean_xml_text(entry.find("cargo")) or "0",
            "dockingspace": clean_xml_text(entry.find("dockingspace")) or "0",
            "integrationspaces": clean_xml_text(entry.find("integrationspaces")) or "0",
            
            # Performance Specs[cite: 13]
            "handling": clean_xml_text(entry.find("handling")) or "0",
            "bp_max": clean_xml_text(entry.find("BodyPointsMax")) or "0",
            "speed_top": clean_xml_text(entry.find("topspeed")) or "0",
            "speed_cruise": clean_xml_text(entry.find("cruise")) or "0",
            "speed_accel": clean_xml_text(entry.find("acceleration")) or "0",
            "speed_decel": clean_xml_text(entry.find("deceleration")) or "0",
            
            # Text Fields[cite: 13]
            "primary_function": clean_xml_text(entry.find("primaryfunction")) or "N/A",
            
            # Arrays & Raw references[cite: 13]
            "components": components_array,
            "raw_node": entry
        }
        
        vehicles_list.append(vehicle_data)

    print(f"[DEBUG END: VEHICLE PARSER] Finished processing. Total successfully packaged: {len(vehicles_list)} vehicles.")
    return sorted(vehicles_list, key=lambda x: x["name"].lower())