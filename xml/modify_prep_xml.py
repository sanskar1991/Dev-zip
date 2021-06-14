from lxml import etree, objectify


metadata = 'presentation.xml'
parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse(metadata, parser)
root = tree.getroot()

print("ROOT: ", root)
print("ROOT.TAG: ", root.tag)

slides = [2, 4, 6]
first_slide_id = 8
slides_id = ['rId'+str(first_slide_id+i-1) for i in slides]
print("LIST: ", slides_id)
tot_slides = 28

for relation in root:
    for i in relation:
        # print("II: ", i.attrib.values())
        try:
            rId = int(i.attrib.values()[-1].split('Id')[-1])
            # rId = int(str_rId)
            # print("rID: ", rId)
            if rId>=first_slide_id and rId<(first_slide_id+tot_slides):
                # print("GGG")
                if 'rId'+str(rId) not in slides_id:
                    # print("IDZZ: ", 'rId'+str(rId))
                    relation.remove(i)
        except:
            pass
            
        
tree.write('output1.xml', pretty_print=True, xml_declaration=True, encoding='UTF-8')
        