# import json module and xmltodict
# module provided by python
import json
import xmltodict

# Converting Python Dictionary to 
# XML and saving to a file 
from dicttoxml import dicttoxml
from xml.dom.minidom import parseString
  
  
# open the input xml file and read
# data in form of python dictionary 
# using xmltodict module
with open("presentation.xml.rels") as xml_file:
    data_dict = xmltodict.parse(xml_file.read())
    xml_file.close()
      
    # generate the object using json.dumps() 
    # corresponding to json data
      
json_data = json.dumps(data_dict)
dict_data = json.loads(json.dumps(data_dict))
# print("TYPE: ", (json_data))
# print("NEW_DATA: ", data_dict, "\nTYTYTY: ", type(dict_data))
      
    # Write the json data to output 
    # json file
with open("data.json", "w") as json_file:
    json_file.write(json_data)
    json_file.close()
    
with open('data.json') as json_file:
    data = json.load(json_file)
    
print("DATAAA: ", data)
print("TYIPOUU: ", type(data))
        
# -----------------
# Variable name of Dictionary is data
xml = dicttoxml(data, attr_type = False)
  
# Obtain decode string by decode()
# function
xml_decode = xml.decode()
  
xmlfile = open("dict.xml", "w")
xmlfile.write(xml_decode)
xmlfile.close()