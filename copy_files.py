# Functions which copies all the relationship files

import shutil
import os
import xmltodict


def copy_rel(src, des):
    """
    copy all relelationship files
    """
    for x in os.walk(src):
        folder = x[0].split('ppt')[1]
        # print("FOLDER: ", folder)
        if folder and '_rels' in folder and 'slides' not in folder:
            # print("SRC: ", src+folder, "\nDES: ", des+folder)
            if os.path.exists(des+folder):
                shutil.rmtree(des+folder)
            shutil.copytree(src+folder, des+folder)
    
    # remove empty directories from output dir
    for dir in os.walk(des):
        if not dir[2]:
            if os.path.exists(dir[0]):
                shutil.rmtree(dir[0])
    print("COPY COMPLETED: ")
    
    return

  
def xml_to_dict(path):
    """
    convert xml to dict
    """
    with open(path) as xml_file:
        data_dict = xmltodict.parse(xml_file.read())
        xml_file.close()
    if isinstance(data_dict["Relationships"]["Relationship"], list):
        data = sorted(data_dict["Relationships"]["Relationship"], key=lambda item: int(item['@Id'].split('Id')[1]))
    else:
        data = [data_dict["Relationships"]["Relationship"]]
    return data


def copy_mandatory(src, des):
    """
    copy mandatory files
    """
    print("COPY MANDATORY CALLING")
    lis = ['slideLayouts', 'theme']
    for dir in lis:
        if os.path.exists(des+dir):
            shutil.rmtree(des+dir)
        shutil.copytree(src+dir, des+dir)
    print("MANDATORY DONE!!")
    
    return





a = ['../customXml/item1.xml', '../customXml/item2.xml', '../customXml/item3.xml', 'slideMasters/slideMaster1.xml', 'slideMasters/slideMaster2.xml', 'slideMasters/slideMaster3.xml', 'slideMasters/slideMaster4.xml', 'notesMasters/notesMaster1.xml', 'handoutMasters/handoutMaster1.xml', 'commentAuthors.xml', 'presProps.xml', 'viewProps.xml', 'theme/theme1.xml', 'tableStyles.xml', 'changesInfos/changesInfo1.xml', 'revisionInfo.xml', 'slides/slide2.xml', '../slideLayouts/slideLayout11.xml', '../slideMasters/slideMaster1.xml', '../slideLayouts/slideLayout1.xml', '../slideLayouts/slideLayout2.xml', '../media/image3.emf', '../slideLayouts/slideLayout3.xml', '../slideLayouts/slideLayout4.xml', '../slideLayouts/slideLayout5.xml', '../slideLayouts/slideLayout6.xml', '../slideLayouts/slideLayout7.xml', '../media/image4.jpeg', '../media/image5.png', '../slideLayouts/slideLayout8.xml', '../slideLayouts/slideLayout9.xml', '../slideLayouts/slideLayout10.xml', '../theme/theme1.xml', '../media/image1.emf', '../media/image2.png', '../media/image8.png', 'slides/slide4.xml', 'slides/slide6.xml', '../notesSlides/notesSlide2.xml', '../notesMasters/notesMaster1.xml', '../theme/theme5.xml', '../slides/slide6.xml', '../media/image23.png', '../media/image24.png', '../media/image25.png']
b = ['../customXml/item1.xml', '../customXml/item2.xml', '../customXml/item3.xml', 'slideMasters/slideMaster1.xml', 'slideMasters/slideMaster2.xml', 'slideMasters/slideMaster3.xml', 'slideMasters/slideMaster4.xml', 'notesMasters/notesMaster1.xml', 'handoutMasters/handoutMaster1.xml', 'commentAuthors.xml', 'presProps.xml', 'viewProps.xml', 'theme/theme1.xml', 'tableStyles.xml', 'changesInfos/changesInfo1.xml', 'revisionInfo.xml', 'slides/slide2.xml', '../slideLayouts/slideLayout11.xml', '../slideMasters/slideMaster1.xml', '../slideLayouts/slideLayout1.xml', '../slideLayouts/slideLayout2.xml', '../media/image3.emf', '../slideLayouts/slideLayout3.xml', '../slideLayouts/slideLayout4.xml', '../slideLayouts/slideLayout5.xml', '../slideLayouts/slideLayout6.xml', '../slideLayouts/slideLayout7.xml', '../media/image4.jpeg', '../media/image5.png', '../slideLayouts/slideLayout8.xml', '../slideLayouts/slideLayout9.xml', '../slideLayouts/slideLayout10.xml', '../theme/theme1.xml', '../media/image1.emf', '../media/image2.png', '../media/image8.png', 'slides/slide4.xml', 'slides/slide6.xml', '../notesSlides/notesSlide2.xml', '../notesMasters/notesMaster1.xml', '../theme/theme5.xml', '../slides/slide6.xml', '../media/image23.png', '../media/image24.png', '../media/image25.png']

lis = []
lit = []
for i in a:
    if i not in b:
        lis.append(i)
        
for i in b:
    if i not in a:
        lit.append(i)
        
print("a: ", lis)
print("b: ", lit)