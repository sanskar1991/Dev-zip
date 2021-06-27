
from lxml import etree

parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse('presentation.xml.rels', parser)
root = tree.getroot()
xml_str = etree.tostring(root, encoding='unicode')
print(xml_str.find('s'))