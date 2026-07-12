import zipfile
import xml.etree.ElementTree as ET

def clean_xml_text(element):
    return element.text.strip() if element is not None and element.text else ""

def get_xml_root(mod_filename):
    """Opens the .mod file and returns the XML root and open zip reference."""
    mod_zip = zipfile.ZipFile(mod_filename, 'r')
    xml_file = mod_zip.open('db.xml')
    return ET.fromstring(xml_file.read()), mod_zip

# from collections import defaultdict