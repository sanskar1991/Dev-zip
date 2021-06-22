import zipfile
import os
import shutil
import xmltodict
import json

from pptx import Presentation
from lxml import etree, objectify
from collections import OrderedDict
from zip_unzip import unzip, zipdir
from copy_files import copy_rel, xml_to_dict, copy_mandatory


def ig_d(dir, files):
    """
    filter ppt folder
    """
    return [f for f in files if f=='ppt']


def ig_f(dir, files):
    """
    filter all files
    """
    return [f for f in files if os.path.isfile(os.path.join(dir, f))]


def new (ft) :
    """
    create, move and unzip the empty output deck
    """
    fq_empty = "resources/Empty.pptx"
    # create
    prs = Presentation()
    prs.save(fq_empty)
    # move
    d = ".".join ([ft, "pptx"]) # "output/41.pptx"
    shutil.move (fq_empty, d)
    # unzip
    unzip (d, d.split('.')[0])
    return


def make_dir(tmp, des, file):
    """
    creates input deck directories which does not exists in the output deck
    """
    for i in os.walk(f'{tmp}/{file}'):
        if not os.path.exists(f'{output_file_loc}/{i[0].split(file)[-1]}'):
            os.makedirs(f'{output_file_loc}/{i[0].split(file)[-1]}')
    return


def first_slide(file_name):
    """
    returns the first slide rId
    """
    prs = Presentation(f'{input_decks}/{file_name}.pptx')
    for i in prs.slides:
        # print("HELLO: ", i, i.slide_id, i.shapes)
        pass
    abc = prs.slides._sldIdLst
    xml = list(abc)
    

def slide_ids(data, slides):
    """
    returns a list of rIds
    """
    # return [for i in data if slide]
    for id in slides:
        slide = f'slide{str(id)}.xml'
        
    # too = []
    #         for i in data:
    #             if slide in i["@Target"]:
    #                 print("111: ", i["@Id"])
    #             pass


def add_files(path, file_name, slides=None):
    """
    returns a list of files that needs to be modified in output deck
    """
    global target_files
    data = xml_to_dict(path)
    if slides:
        global sldIds
        sldIds = []
        # get total slides
        prs = Presentation(base_path+'/presentations/'+file_name+'.pptx')
        tot_slides = len(prs.slides._sldIdLst)
        # get rId of first slide
        first_slide = "slide1.xml"
        first_slide_id = int([i["@Id"] for i in data if first_slide in i['@Target']][0].split('Id')[1])
        
        files = []
        for i in data:
            current_rId = int(i['@Id'].split('Id')[1])
            if (first_slide_id > current_rId) or (current_rId > (first_slide_id+tot_slides-1)):
                files.append(i['@Target'])
        
        target_files = target_files + files
        for id in slides:
            slide = f'slide{str(id)}.xml'
            sldIds.append([i["@Id"] for i in data if slide in i["@Target"] and "http" not in i["@Target"]][0])
            target_files.append([i["@Target"] for i in data if slide in i["@Target"] and "http" not in i["@Target"]][0])
            shutil.copy(tmp_path+'/'+file_name+"/ppt/slides/_rels/"+slide+".rels", output_path+'/'+str(render_id)+'/ppt/slides/_rels/')
            add_files(tmp_path+'/'+file_name+"/ppt/slides/_rels/"+slide+".rels",file_name)
    else:
        for i in data:
            if i["@Target"] in target_files:
                pass
            elif "http" not in i["@Target"]:
                # print("This time: ", i["@Target"])
                target_files.append(i['@Target'])
                if ".." in i['@Target'] and "xml" in i['@Target']:
                    path = tmp_path+'/'+file_name+"/ppt/"+i['@Target'].split('..')[1].split('/')[1]+"/_rels/"+i['@Target'].split('..')[1].split('/')[2]+".rels"
                    if os.path.exists(path):
                        add_files(path, file_name)
    
    # copy files from tmp dir to output dir
    # for i in target_files:
    #     if '../' in i:
    #         if os.path.exists(tmp_path+'/'+file_name+'/ppt/'+i[3:]):
    #             shutil.copy(tmp_path+'/'+file_name+'/ppt/'+i[3:], output_path+'/'+str(render_id)+'/ppt/'+i[3:].split('/')[0])
    #     else:
    #         shutil.copy(tmp_path+'/'+file_name+'/ppt/'+i, output_path+'/'+str(render_id)+'/ppt/'+i.split('/')[0])
    return target_files, sldIds
    
def select_all():
    """
    if slides is None
    copy all the file content from input to output deck
    """
    pass


def modify(inp_root, out_root, tag_dict, i_tree, o_tree):
    """
    modify content of files
    """
    for i, o in tag_dict.items():
        # print("TYPE: ", type(i), type(o))
        if o == 0:
            pass
        else:
            subtag1 = o_tree.find(o)
            subtag2 = etree.Element(i)
            subtag1.addnext(subtag2)
            
    with open (f'{output_file_loc}/ppt/presentation.xml', 'wb') as f:
        f.write(etree.tostring(out_root, pretty_print = True))
    
    return
    # subtag1 = i_tree.find("subtag1")
    # subtag2 = etree.Element("subtag2", subattrib2="2")
    # subtext = etree.SubElement(subtag2, "subtext")
    # subtext.text = "text2"
    # subtag1.addnext(subtag2)   # Add subtag2 as a following sibling of subtag1

    # print( etree.tostring(tag, pretty_print=True))


def tag(inp_tag, out_tag):
    """
    returns a dict of tags
    """
    tag_dict = OrderedDict()
    for i in range(len(inp_tag)):
        if inp_tag[i] not in out_tag:
            if i == 0:
                tag_dict[inp_tag[i]] = 0
            else:
                tag_dict[inp_tag[i]] = inp_tag[i-1]
    return tag_dict
    
  

def tree(src, des):
    """
    pass the path of the xml document to enable the parsing process
    """
    parser = etree.XMLParser(remove_blank_text=True)
    inp_tree = etree.parse(src, parser)
    out_tree = etree.parse(des, parser)
    inp_root = inp_tree.getroot()
    out_root = out_tree.getroot()
    
    return inp_root, out_root, inp_tree, out_tree
    


def pre_xml(file_name, slides): # tmp/41/{file_name}/ppt/presentation.xml
    """
    modify presentation.xml file
    """
    src_xml = f'{tmp_path}/{file_name}/ppt/presentation.xml'
    des_xml = f'{output_file_loc}/ppt/presentation.xml'
    
    inp_root, out_root, i_tree, o_tree = tree(src_xml, des_xml)
    
    # for relation in inp_root:
    #     print("FFFF: ", relation.tag, relation.attrib)
    
    inp_tag = [relation.tag for relation in inp_root]
    out_tag = [relation.tag for relation in out_root]
    # print("INP: ", inp_tag, "\nOUT: ", out_tag)
    
    tag_dict = tag(inp_tag, out_tag)
    
    if slides:
        # slide_1 = if_slides()
        pass
    print("TAGAGA: ", tag_dict)
    for relation in inp_root:
        for ele in relation:
            try:
                pass
                # print("ELE: ", ele, ele.attrib, ele.attrib.values())
            except:
                pass
            # try:
            #     rId = int(ele.attrib.values()[-1].split('Id')[-1])
            #     if rId>=first_slide_id and rId<(first_slide_id+tot_slides):
            #         # print("GGG")
            #         if 'rId'+str(rId) not in slides_id:
            #             relation.remove(ele)
            # except:
            #     pass
                
    modify(inp_root, out_root, tag_dict, i_tree, o_tree)
    # print("LIS: ", lis)
    return
        
        



def deck_handle(id, msg):
    """
    handle the deck and select files for output deck
    """
    file_name, slides = msg['d'], msg['s']
    new(output_file_loc)

    # unzip the input deck
    unzip(f'{input_decks}/{file_name}.pptx', f'{tmp_path}/{file_name}')
    
    make_dir(tmp_path, output_file_loc, file_name)
    prep_xml_path = f'{tmp_path}/{file_name}/ppt/_rels/presentation.xml.rels'
    a, b = add_files(prep_xml_path, file_name, slides)
    print("TARGET11: ", a, b)
    pre_xml(file_name, slides)
    
    return
    
    
    # if os.path.isdir(f'{output_path}/{str(render_id)}'):
    #     # copy all the necessary files with folder architecture
    #     shutil.copytree(f'{tmp_path}/{file_name}', f'{output_path}/{str(render_id)}', ignore=ig_d)
    #     shutil.copytree(f'{tmp_path}/{file_name}/ppt', f'{output_path}/{str(render_id)}/ppt', ignore=ig_f)
    
    if slides:
        pass
    # add_files()
    else:
        pass
        
    
    
if __name__ == '__main__':
    
    base_path = os.path.dirname(os.path.realpath(__file__))
    print("CURRENT_DIR:", base_path)
    target_files = []
    
    # load the message
    # file = open('sample_input.json')
    # sample_msg = json.load(file)
    # file.close()
    sample_msg = [41,{'d': 'Onboarding','s':  [2,4,6]}]
    # sample_msg = [41,{'d': 'Presentation1','s':  [1]}]
    # sample_msg = [41,{'d': 'BI Case Studies','s':  [2, 3]}]

    render_id = sample_msg.pop(0)
    
    output_path = f'{base_path}/output'
    tmp_path = f'{base_path}/tmp/{render_id}'
    input_decks = f'{base_path}/presentations'
    output_file_loc = f'{output_path}/{render_id}'
    
    print("TMP_PATH:", tmp_path, '\nOUT_PATH: ', output_path)   

    try:
        os.makedirs(tmp_path)
        os.makedirs(output_path)
    except:
        print("DIR ALREADY EXIST")
    
    # iterating all the messages
    while sample_msg:
        deck_handle(render_id, sample_msg.pop(0))

