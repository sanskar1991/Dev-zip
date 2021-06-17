# Functions which copies all the relationship files

import shutil
import os
import xmltodict


# copy all rel files
def copy_rel(src, des):
    for x in os.walk(src):
        folder = x[0].split('ppt')[1]
        # print("FOLDER: ", folder)
        if folder and '_rels' in folder and 'slides' not in folder:
            # print("SRC: ", src+folder, "\nDES: ", des+folder)
            if os.path.exists(des+folder):
                shutil.rmtree(des+folder)
            shutil.copytree(src+folder, des+folder)
    # remove empty directories
    for dir in os.walk(des):
        if not dir[2]:
            if os.path.exists(dir[0]):
                shutil.rmtree(dir[0])
        
    print("COPY COMPLETED: ")

  
# convert xml to dict
def xml_to_dict(path):
    with open(path) as xml_file:
        data_dict = xmltodict.parse(xml_file.read())
        xml_file.close()
    if isinstance(data_dict["Relationships"]["Relationship"], list):
        data = sorted(data_dict["Relationships"]["Relationship"], key=lambda item: int(item['@Id'].split('Id')[1]))
    else:
        data = [data_dict["Relationships"]["Relationship"]]
    return data


# copy mandatory files
def copy_mandatory(src, des):
    print("COPY MANDATORY CALLING")
    lis = ['slideLayouts', 'theme']
    for dir in lis:
        if os.path.exists(des+dir):
            shutil.rmtree(des+dir)
        shutil.copytree(src+dir, des+dir)
    print("MANDATORY DONE!!")