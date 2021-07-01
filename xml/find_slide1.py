from lxml import etree
# from collections import OrderedDict
# import json

parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse('tableStyles.xml', parser)
root = tree.getroot()

# print("OOO: ", root.__dir__())
# print()
# print("NSMP: ", root.nsmap)
# print("TYPE: ", type(root.nsmap))
# nmsps =  root.nsmap['r']
# print("NN: ", nmsps)
print("ROOT: ", root.tag)
print("ATTRIB: ", root.attrib)
for relation in root:
    print("FF: ", relation.tag)
# print("ATTRIB: ", len(root.attrib['def']))
