# story_parser.py
import xml.etree.ElementTree as ET
from fg_parser import clean_xml_text

def harvest_ordered_story_pages(root):
    """
    Parses the refmanualindex by strictly sorting chapters, subchapters,
    and refpages by their internal <order> numerical tags.
    """
    elements = []
    ref_index = root.find('.//reference/refmanualindex')
    ref_data = root.find('.//reference/refmanualdata')
    
    if ref_index is None or ref_data is None:
        return elements

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

    for ch_order, ch_name, chapter_node in ordered_chapters:
        if ch_name:
            elements.append(('chapter', ch_name))
        
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
        
        for subch_order, subch_name, subch_node in ordered_subchapters:
            if subch_name:
                elements.append(('subchapter', subch_name))
            
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
            
            for pg_order, pg_name, pg_node in ordered_refpages:
                record_node = pg_node.find('.//recordname')
                if record_node is not None and record_node.text:
                    page_id = record_node.text.split('.')[-1]
                    page_data_node = ref_data.find(page_id)
                    
                    if page_data_node is not None and pg_name:
                        if pg_name.lower() == "npcs":
                            continue
                        elements.append(('page', pg_name, page_data_node))
                        
    return elements