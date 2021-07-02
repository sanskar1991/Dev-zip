from lxml import etree
# from collections import OrderedDict
# import json

# parser = etree.XMLParser(remove_blank_text=True)
# tree1 = etree.parse('inp_tableStyles.xml', parser)
# root1 = tree1.getroot()
# tree2 = etree.parse('out_tableStyles.xml', parser)
# root2 = tree2.getroot()

# tree = etree.parse('a.xml', parser)
# root = tree.getroot()

# tree3 = etree.parse('b.xml', parser)
# root3 = tree3.getroot()

# for relation in root1:
    
#     print("TAG: ", relation.tag)
#     root2.append()
    
import lxml.etree as ET
filename = "inp_tableStyles.xml"
appendtoxml = "out_tableStyles.xml"
output_file = appendtoxml.replace('.xml', '') + "_editedbyed.xml"

parser = ET.XMLParser(remove_blank_text=True)
tree = ET.parse(filename, parser)
root = tree.getroot()

# for i in root:
#     print("TAG: ", i.tag)
nmsps = root.nsmap['a']
out_tree = ET.parse(appendtoxml, parser)
out_root = out_tree.getroot()

print("TTT: ", root[0].tag)

for i in root:
    print("II: ", i)
    print("LL: ", i.tag)
for path in [f"{root[0].tag}"]:
# for path in [f".//{{{nmsps}}}tblStyle"]:
    for elt in root.findall(path):
        out_root.append(elt)

out_tree.write(output_file, pretty_print=True, xml_declaration=True, encoding='UTF-8')
# import os
# print("HHH: ", os.listdir('..'))
# print()
# a = tree.find('process')
# print("AA: ", a)

# b = tree3.find('process')
# b.addnext(a)
    
# # print()
# # print(root.__dir__())
# # print()
# # root.append(a)
# # root.addnext(a)
# # print("AA: ", a)
# tree3.write('c.xml', pretty_print=True, xml_declaration=True, encoding='UTF-8')
    