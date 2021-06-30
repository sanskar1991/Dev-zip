from lxml import etree


parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse('presentation.xml', parser)
root = tree.getroot()

for relation in root:
    print("TAG: ", relation.tag)




# tree.write('demo.xml.rels', pretty_print=True, xml_declaration=True, encoding='UTF-8')
    