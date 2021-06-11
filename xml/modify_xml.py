# importing element tree
# under the alias of ET
# import xml.etree.ElementTree as ET
from lxml import etree, objectify

# Passing the path of the
# xml document to enable the
# parsing process
metadata = 'presentation.xml.rels'
parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse(metadata, parser)
root = tree.getroot()
  
# getting the parent tag of
# the xml document
root = tree.getroot()
  
# printing the root (parent) tag
# of the xml document, along with
# its memory location
print("ROOT: ", root)
print("ROOT.TAG: ", root.tag)
  
# printing the attributes of the
# first tag from the parent

# print(root.findall(''))
slides = [1]
first_slide_id = 2
slides_id = ['rId'+str(first_slide_id+i-1) for i in slides]
print("LIST: ", slides_id)
tot_slides = 2

for relation in root:
    attrib = relation.attrib
    # print("REL_ATTR: ", attrib, type(attrib))
    print("IDDDDDD: ", attrib.get('Id'))
    
    if int(attrib.get('Id').split('Id')[1]) >= first_slide_id and int(attrib.get('Id').split('Id')[1])<(first_slide_id+tot_slides):
        if attrib.get('Id') not in slides_id:
            root.remove(relation)
            
    # if relation.attrib.get('Id') not in slides:
        # root.remove(relation)
# print(root.attrib)

tree.write('output1.xml.rels', pretty_print=True, xml_declaration=True, encoding='UTF-8')
  
# printing the text contained within
# first subtag of the 5th tag from
# the parent
# print(root[5][0].text)