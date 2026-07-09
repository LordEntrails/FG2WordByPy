import zipfile
import xml.etree.ElementTree as ET

mod_filename = "FS_MtN.mod"

def clean_xml_text(element):
    """Safely extracts raw text from an element."""
    return element.text.strip() if element is not None and element.text else ""

def stringify_formatted_text(node):
    """
    Recursively converts a <formattedtext> block into plain text with 
    simple markdown hints so we know how to style it in Word later.
    """
    if node is None:
        return ""
    
    text_pieces = []
    
    # Handle text immediately inside the parent tag
    if node.text:
        text_pieces.append(node.text)
        
    for child in node:
        # Check block levels or structural flags
        if child.tag == 'h':
            text_pieces.append(f"\n[HEADING: {stringify_formatted_text(child)}]\n")
        elif child.tag == 'frame':
            text_pieces.append(f"\n[CALLOUT/BOXTEXT: {stringify_formatted_text(child)}]\n")
        elif child.tag == 'li':
            text_pieces.append(f"\n* {stringify_formatted_text(child)}")
        elif child.tag == 'p':
            # Recurse inline formats inside a paragraph (like bold/italic)
            p_text = stringify_formatted_text(child)
            if p_text:
                text_pieces.append(f"\n{p_text}\n")
        elif child.tag == 'b':
            text_pieces.append(f"**{stringify_formatted_text(child)}**")
        elif child.tag == 'i':
            text_pieces.append(f"*{stringify_formatted_text(child)}*")
        else:
            # Fallback for any other inline text nesting
            text_pieces.append(stringify_formatted_text(child))
            
        # Grab text that immediately trails a child element tag
        if child.tail:
            text_pieces.append(child.tail)
            
    return "".join(text_pieces).strip()

try:
    with zipfile.ZipFile(mod_filename, 'r') as mod_zip:
        with mod_zip.open('db.xml') as xml_file:
            xml_data = xml_file.read()
            
    root = ET.fromstring(xml_data)
    print("--- Successfully loaded XML for Story Parsing! ---\n")
    
    # ==========================================
    # 1. HARVEST BASIC STORIES (<encounter>)
    # ==========================================
    print("=== PROCESSING BASIC STORIES ===")
    basic_stories = root.find('./encounter')
    if basic_stories is not None:
        for entry in basic_stories:
            title = clean_xml_text(entry.find("name[@type='string']"))
            if title:
                print(f"\nStory Title: {title}")
                print("-" * 30)
                text_node = entry.find("text[@type='formattedtext']")
                if text_node is not None:
                    print(stringify_formatted_text(text_node))
                print("=" * 50)

    print("\n" + "#" * 60 + "\n")

    # ==========================================
    # 2. HARVEST ADVANCED MANUALS (<refmanualdata>)
    # ==========================================
    print("=== PROCESSING ADVANCED REFERENCE MANUAL PAGES ===")
    ref_data = root.find('.//reference/refmanualdata')
    
    if ref_data is not None:
        for page in ref_data:
            page_name = clean_xml_text(page.find("name[@type='string']"))
            if not page_name:
                continue
                
            print(f"\nManual Page: {page_name.upper()}")
            print("=" * 50)
            
            blocks_node = page.find('blocks')
            if blocks_node is not None:
                # Harvest blocks and sort them strictly by their internal 'order' number tag
                ordered_blocks = []
                for block in blocks_node:
                    order_val = block.find("order[@type='number']")
                    order_num = int(order_val.text) if order_val is not None and order_val.text else 999
                    ordered_blocks.append((order_num, block))
                
                # Sort the list using the order number
                ordered_blocks.sort(key=lambda x: x[0])
                
                # Print the content in the exact order it displays in Fantasy Grounds
                for order_num, block in ordered_blocks:
                    block_type = clean_xml_text(block.find("blocktype[@type='string']"))
                    print(f"--- Block (Order: {order_num}, Type: {block_type}) ---")
                    
                    # Look for embedded text blocks
                    text_node = block.find("text[@type='formattedtext']")
                    if text_node is not None:
                        print(stringify_formatted_text(text_node))
                        
                    # Detect embedded images / asset tokens
                    image_node = block.find(".//image")
                    if image_node is not None:
                        bitmap_node = image_node.find(".//bitmap")
                        if bitmap_node is not None and bitmap_node.text:
                            print(f"[EMBEDDED IMAGE ASSET: {bitmap_node.text}]")
                    print()
            print("#" * 50)

except FileNotFoundError:
    print(f"Error: Could not find the module file '{mod_filename}'.")