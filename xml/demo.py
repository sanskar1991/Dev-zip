from lxml import etree
from collections import OrderedDict
import json

parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse('presentation.xml', parser)
root = tree.getroot()

# print("OOO: ", root.__dir__())
# print()
# print("NSMP: ", root.nsmap)
# print("TYPE: ", type(root.nsmap))
nmsps =  root.nsmap['r']
print("NN: ", nmsps)

d1 = OrderedDict()
tag_list = ['{http://schemas.openxmlformats.org/presentationml/2006/main}notesMasterIdLst', '{http://schemas.openxmlformats.org/presentationml/2006/main}handoutMasterIdLst', '{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst', '{http://schemas.openxmlformats.org/presentationml/2006/main}extLst']


for relation in root:
    if relation.tag in tag_list:
        for ele in relation:
            attrib = ele.attrib
            # print("ATTRIB: ", attrib)
            tag = ele.tag
            # if attrib.get(f"{{{nmsps}}}id"):
            #     if attrib.get('Id'):
            #         d1[relation.tag] = [tag, attrib['Id'], attrib.get(f"{{{nmsps}}}id")]
            #     pass
            if relation.tag in d1:
                print("IF : ", relation.tag)
                try:
                    val = d1[relation.tag]
                    val.append([tag, attrib.get('id'), attrib.get(f"{{{nmsps}}}id")])
                    d1[relation.tag] = val
                except:
                    pass
            else:
                print("ELSE: ", relation.tag)
                d1[relation.tag] = [[tag, attrib.get('id'), attrib.get(f"{{{nmsps}}}id")]]
obj = json.dumps(d1)
with open("new.json", "w") as outfile:
        outfile.write(obj)
print("DDDDD: ", d1)
        
        # try:
        #     tag = ele.tag
        #     rId = attrib[f"{{{nmsps}}}id"]
        #     d1[rId] = [tag]
        # except:
        #     pass
        # # print("AAA: ", attrib)
        # b = attrib[f"{{{nmsps}}}id"]
        
        


# tree.write('demo.xml.rels', pretty_print=True, xml_declaration=True, encoding='UTF-8')
    