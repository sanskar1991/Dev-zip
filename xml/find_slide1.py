from pptx import Presentation
from lxml import etree, objectify


parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse('presentation.xml.rels', parser)
root = tree.getroot()

slide = 'slide1.xml'
for relation in root:
    attrib = relation.attrib
    print("NEW: ", attrib)
    if slide in attrib['Target']:
        print("RID: ", attrib['Id'])
    