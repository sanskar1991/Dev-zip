import xml.etree.ElementTree as ET

file = ET.parse('presentation.xml')
root = file.getroot()
print("FFFFF: ", root.findall('.//xml'))
for elem in root.findall('.//trigger/external'):
    check_req_elems = elem.find('./action[@name="check_req"]')
    check_elem = elem.find('./action[@name="ckeck"]')
    print("CHECKKK: ", check_req_elems)
    if check_req_elems is not None and check_elem is not None:
            elem.remove(check_elem)

file.write('b.xml')