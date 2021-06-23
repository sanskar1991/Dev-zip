import os
import shutil

from pptx import Presentation
from lxml import etree, objectify
# from base import first_slide

def copy_mandatory(src, des):
    """
    copy mandatory files
    """
    print("COPY MANDATORY CALLING")
    lis = ['slideLayouts', 'theme']
    for dir in lis:
        if os.path.exists(f'{des}/{dir}'):
            shutil.rmtree(f'{des}/{dir}')
        shutil.copytree(f'{src}/{dir}', f'{des}/{dir}')
    print("MANDATORY DONE!!")
    
    return


def copy_prep_xml(path):
    """
    copy main relationship and xml file of the deck
    """
    print("COPY_PREP_XML CALLIMG...")
    # global tot_slides, slides
    slide_1 = first_slide(path)
    slides_id = ['rId'+str(first_slide_id+i-1) for i in slides]
    
    # Setting up the paths for xml and rels file
    rels_path = path+'_rels/presentation.xml.rels'
    xml_path = path+'presentation.xml'
    
    # Passing the path of the xml document to enable the parsing process
    # for rels file
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(rels_path, parser)
    root = tree.getroot()

    # iterating root
    for relation in root:
        attrib = relation.attrib

        if int(attrib.get('Id').split('Id')[1]) >= first_slide_id and int(attrib.get('Id').split('Id')[1])<(first_slide_id+tot_slides):
            if attrib.get('Id') not in slides_id:
                root.remove(relation)
    tree.write(output_path+'/'+str(render_id)+'/ppt/_rels/presentation.xml.rels', pretty_print=True, xml_declaration=True, encoding='UTF-8')
    
    # Passing the path of the xml document to enable the parsing process
    # for XML file
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(xml_path, parser)
    root = tree.getroot()
    for relation in root:
        for ele in relation:
            try:
                rId = int(ele.attrib.values()[-1].split('Id')[-1])
                if rId>=first_slide_id and rId<(first_slide_id+tot_slides):
                    # print("GGG")
                    if 'rId'+str(rId) not in slides_id:
                        relation.remove(ele)
            except:
                pass
        tree.write(output_path+'/'+str(render_id)+'/ppt/presentation.xml', pretty_print=True, xml_declaration=True, encoding='UTF-8')
    print("COMPLETED!!1")



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
    copy_mandatory(src, des)
    copy_prep_xml(src)
    return