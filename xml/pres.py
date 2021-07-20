from lxml import etree, objectify

metadata = 'presProps.xml'
parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse(metadata, parser)
root = tree.getroot()

a = {}

nm = root.nsmap['p']
print("NNN: ", f"{{{nm}}}showPr")

for relation in [root[0].tag]:
    # for ele in relation:
        print("AAA: ", relation)
        # print("YYY: ", type(relation))
        fp = root.find(f"{{{nm}}}extLst")
        # print("FFF: ", fp)
        for ele in fp:
            attrib = ele.attrib
            if attrib['uri'] not in a.keys():
                a[attrib['uri']] = ele
            # print("KKK: ", att)
            # print("JJJ: ", ele.tag)
        # for ele in root.findall(relation):
        #     print("JJJ: ", ele.tag)
        # print("AA: ", a)
