import os
import shutil
import json
import xmltodict
import natsort
import zipfile
import re
import pathlib

from pptx import Presentation
from lxml import etree
from collections import OrderedDict


def unzip(src, des):
    """
    unzip the deck
    """
    with zipfile.ZipFile(src, 'r') as zip_ref:
        zip_ref.extractall(des)
    return


def zipdir(path, file_name):
    """
    zip extracted deck to get output deck
    """
    length = len(path)
    zipf = zipfile.ZipFile('output/'+f'Test_{file_name}.pptx', 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(path):
        folder = root[length:] # path without "parent"
        for file in files:
            zipf.write(os.path.join(root, file), os.path.join(folder, file))
    zipf.close()
    return


def gen_tree(path):
    """
    pass the path of the xml document to enable the parsing process
    """
    print("CALLING.. Tree")
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    root = tree.getroot()    
    return root, tree


def max_rId():
    """
    returns maximum rId
    """
    path = f'{output_path}/ppt/_rels/presentation.xml.rels'
    root, tree = gen_tree(path)
    
    rIds = []
    
    for relation in root:
        attrib = relation.attrib
        rId = int(attrib.get('Id').split('Id')[-1])
        rIds.append(rId)
    return {'rId': max(rIds)}


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


def total_slides(path):
    """
    returns total number of slides
    """
    print("CALLING.. total_slides")
    prs = Presentation(path)
    tot_slides = len(prs.slides._sldIdLst)
    return tot_slides


def first_slide(path):
    """
    returns the first slide rId
    """
    print("CALLING.. First_slide")
    root, _ = gen_tree(path)

    slide = 'slide1.xml'
    for relation in root:
        attrib = relation.attrib
        if slide in attrib['Target']:
            return int(attrib['Id'].split('Id')[-1])


def get_rels(dict_1):
    """
    list latest .rels files
    """
    # a = dict_1.keys()
    a = [i for i in dict_1.values()]
    lis = natsort.natsorted([i for i in a if '_rels' in i])
    return lis


def new(path):
    """
    create, move and unzip the empty output deck
    """
    fq_empty = "resources/Empty.pptx"
    # create
    prs = Presentation()
    prs.save(fq_empty)
    # move
    d = ".".join ([path, "pptx"]) # "output/41.pptx"
    shutil.move (fq_empty, d)
    # unzip
    unzip (d, d.split('.')[0])
    m_rId = max_rId()
    return m_rId


def make_dir(file_name): # output_file_loc = des
    """
    creates input deck directories which does not exists in the output deck
    """
    for i in os.walk(f'{tmp_path}/{file_name}'):
        fld = i[0].split(file_name)[-1]
        if fld:
            loc = f"{output_path}{fld}"
            if not os.path.exists(f'{output_path}/{fld}'):
                os.makedirs(f'{output_path}/{fld}')
    print("MAKE_DIR completed...")        
    return


def make_structure(file_name):
    """
    creates structure of input deck
    """
    for i in os.walk(f'{tmp_path}/{file_name}'):
        fld = i[0].split(file_name)[-1]
        if fld:
            loc = f"{output_path}{fld}"
            if 'ppt' not in loc and (file_name not in loc):
                shutil.rmtree(f'{output_path}/{fld}')
                shutil.copytree(f'{tmp_path}/{file_name}/{i[0].split(file_name)[-1]}', f'{output_path}/{fld}')
            elif 'ppt' in loc:
                if 'slideLayouts' in loc or 'slideMasters' in loc or 'theme' in loc:
                    # copy_mandatory()
                    pass
            
            # if not os.path.exists(f'{output_file_loc}/{fld}'):
            #     if 'ppt' not in f'{output_file_loc}/{fld}':
            #         shutil.copytree(f"{tmp_path}/{file}/{i[0].split(f'{file}/')[-1]}", f'{output_file_loc}/{i[0].split(file)[-1]}')
            #     else:
            #         os.makedirs(f'{output_file_loc}/{i[0].split(file)[-1]}')
    return


def add_files(path, file_name, target_files, slides=None):
    """
    returns a list of files that needs to be modified in output deck
    """
    print("CALLING.. Add_files")
    data = xml_to_dict(path)
    if slides:
        global sldIds
        
        tot_slides = total_slides(f'{input_decks}/{file_name}.pptx')
        first_slide_id = first_slide(path)
        
        files = []
        for i in data:
            current_rId = int(i['@Id'].split('Id')[1])
            if (first_slide_id > current_rId) or (current_rId > (first_slide_id+tot_slides-1)):
                if 'slideLayouts' not in i['@Target'] and 'slideMasters' not in i['@Target'] and 'theme' not in i['@Target']:
                    files.append(i['@Target'])
        for i in files:
            if '/' in i:
                a = i.split('/')
                fld, fl = a[0], a[1]
                if os.path.exists(f'{tmp_path}/{file_name}/ppt/{fld}/_rels'):
                    if os.path.isfile(f'{tmp_path}/{file_name}/ppt/{fld}/_rels/{fl}.rels'):
                        target_files.append(f'{fld}/_rels/{fl}.rels')
                    # fl_list = os.listdir(f'{tmp_path}/{file_name}/ppt/{fld}/_rels')
                    # for i in fl_list:
                    #     if 
                    
            pass# if rel exists then will add it to list
        
        target_files = target_files + files
        
        for id in slides:
            slide = f'slide{str(id)}.xml'
            sldIds.append([i["@Id"] for i in data if slide in i["@Target"] and "http" not in i["@Target"]][0])
            target_files.append([i["@Target"] for i in data if slide in i["@Target"] and "http" not in i["@Target"]][0])
            target_files.append(f'slides/_rels/{slide}.rels')
            add_files(f'{tmp_path}/{file_name}/ppt/slides/_rels/{slide}.rels', file_name, target_files)
    else:
        for i in data:
            # handling duplicacy
            if i["@Target"] in target_files or i["@Target"][3:] in target_files:
                pass
            elif "http" not in i["@Target"]:
                # if 'slideLayouts' not in i['@Target'] and 'slideMasters' not in i['@Target'] and 'theme' not in i['@Target']:
                if '../' in i["@Target"]:
                    target_files.append(i['@Target'][3:])
                else:
                    target_files.append(i['@Target'])
                
                if ".." in i['@Target'] and "xml" in i['@Target']:
                    path = f"{tmp_path}/{file_name}/ppt/{i['@Target'].split('..')[1].split('/')[1]}/_rels/{i['@Target'].split('..')[1].split('/')[2]}.rels"
                    
                    if os.path.exists(path):
                        # handling rels files
                        target_files.append(path.split('ppt/')[1])
                        add_files(path, file_name, target_files)
                        
        # print("UUU2: ", target_files)
    # print("UUU: ", target_files)
    return target_files


def get_fld_fl(file):
    """
    returns folder anme and file name
    """
    if '_rels' in file: # slides/_rels/slide2.xml.rels
        sp = file.split('/')
        fl_name = sp[-1]
        fld_name = f'{sp[0]}/{sp[1]}'
    elif '../' in file:
        _,fld_name,fl_name = file.split('/')
    else:
        fld_name,fl_name = file.split('/')
    
    return fld_name, fl_name


def list_target(target_files, d2):
    """
    creates a dict with number of files
    """
    # d2 = OrderedDict()
    count = 0
    for file in target_files:
        if '/' in file:
        # get folder and file name
            fld, fl = get_fld_fl(file)
            if fld not in d2:
                d2[fld] = 0
        if 'slideMasters' not in d2:
            d2['slideMasters'] = 0
        if 'slideLayouts' not in d2:
            d2['slideLayouts'] = 0
        if 'theme' not in d2:
            d2['theme'] = 0
    return d2
    

def rename(path, fld, fl, dict_2): # fld=media, fl=image1.png
    """
    rename a file
    """
    d1 = OrderedDict()
    
    ext = ''.join(pathlib.Path(fl).suffixes)
    name = re.findall(r'(\w+?)(\d+)', fl)[0][0]
    
    count = dict_2[fld]+1
    new_name = f'{name}{count}{ext}'
    # if 'slideMasters' in fld:
        # print("NEW_NAME", new_name)
        # print("GGG: ", dict_2)
    if 'ppt' in path:
        shutil.copy(f'{path}/{fld}/{fl}', f"{output_path}/ppt/{fld}/{new_name}")
    else:
        shutil.copy(f'{path}/{fld}/{fl}', f"{output_path}/{fld}/{new_name}")
    d1[f'{fld}/{fl}'] = f'{fld}/{new_name}'
    dict_2[fld] = count
    # print("RENAME: ", fld, fl, new_name)
    return d1


def del_files(rels_fl, last_fl, path):
    """
    delete extra files
    """
    for i in rels_fl:
        if i[:-5] not in last_fl:
            os.remove(f'{path}/{i}')


def copy_mandatory(src, des, deck):
    """
    copy mandatory files
    """
    # print("SRC: ", src, "\nDES: ", des)
    print("COPY MANDATORY CALLING")
    m_list = ['slideLayouts', 'theme', 'slideMasters']
    d1 = OrderedDict()
    if deck == 1:
        for fl in m_list:
            count = 0
            for i in os.walk(f'{src}/ppt/{fl}'):
                count = len(i[2])
            if os.path.exists(f'{des}/{fl}'):
                shutil.rmtree(f'{des}/{fl}')
            shutil.copytree(f'{src}/ppt/{fl}', f'{des}/{fl}')
            dict_2[fl] = count
            if os.path.exists(f'{src}/ppt/{fl}/_rels'):
                dict_2[f'{fl}/_rels'] = count
    else:
        # print("SSSS: ", dict_2)
        for i in m_list:
            if os.path.exists(f'{des}/{i}'):
                for j in os.walk(f'{src}/ppt/{i}'):
                    # j = (path, [folder], [file])
                    fld = j[0].split('ppt/')[1]
                    
                    # print("TTT: ", fld, "--", j[2]) # array
                    fl_list = natsort.natsorted(j[2])
                    for x in fl_list:
                        ext = ''.join(pathlib.Path(x).suffixes)
                        name = re.findall(r'(\w+?)(\d+)', x)[0][0]
                        count = dict_2[fld]+1
                        new_name = f'{name}{count}{ext}'
                        # print("UUUU: ", new_name, count)
                        # print("UUUU: ", f'{src}/ppt/{i}/{x}')
                        # print("1111: ", f'{des}/{fld}/{new_name}')
                        shutil.copy(f'{src}/ppt/{fld}/{x}', f'{des}/{fld}/{new_name}')
                        
                        d1[f'{i}/{x}'] = f'{fld}/{new_name}'
                        dict_2[fld] = count
    
    # remove empty folders
    for i in os.walk(des):
        if not i[2]:
            shutil.rmtree(i[0])

    
    print("MANDATORY DONE!!")
    
    
    return dict_2, d1


def copy_target(target_files, file_name, tmp_loc, dict_2):
    """
    copy target files from tmp to output folder 
    """
    d_1 = OrderedDict()
    
    target_files = natsort.natsorted(target_files)
    for i in target_files:
        if '/' in i:
            if 'slideLayouts' not in i and 'slideMasters' not in i and 'theme' not in i:
                fld_name,fl_name = get_fld_fl(i)
                # print("DDDD: ", fl_name)
                if os.path.exists(f'{tmp_loc}/ppt/{fld_name}/{fl_name}'):
                    path = f'{tmp_loc}/ppt'
                    d_1.update(rename(path, fld_name, fl_name, dict_2))
                else:
                    d_1.update(rename(tmp_loc, fld_name, fl_name, dict_2))
        else:
            if not os.path.isfile(f'{output_path}/ppt/{i}'):
                shutil.copyfile(f'{tmp_loc}/ppt/{i}', f'{output_path}/ppt/{i}')
    
    return d_1


def update_rels(fl_list):
    """
    update latest .rels files
    """


def deck_handler(id, msg, deck, dict_2):
    """
    handle the deck and select files for output deck
    """
    file_name, slides = msg['d'], msg['s']
    target_files = []
    
    m_rId = new(output_path)
    # print("MMM: ", m_rId)
    dict_2.update(m_rId)
    tmp_loc = f'{tmp_path}/{file_name}'
    
    # unzip the input deck
    unzip(f'{input_decks}/{file_name}.pptx', tmp_loc)
    
    # creates folder structure of the input deck
    make_dir(file_name)
    
    prep_xml_path = f'{tmp_loc}/ppt/_rels/presentation.xml.rels'
    
    if deck == 1:
        make_structure(file_name)
        
    # print("DICT_2: ", dict_2)
    target_files = add_files(prep_xml_path, file_name, target_files, slides)
    print("TARGET: ", target_files)
    dict_2.update(list_target(target_files, dict_2))
    # print("YY: ", dict_2)
    dict_1 = copy_target(target_files, file_name, tmp_loc, dict_2)
    a, b = copy_mandatory(tmp_loc, f'{output_path}/ppt', deck)
    dict_2.update(a)
    dict_1.update(b)
    
    obj_1 = json.dumps(dict_1)
    obj_2 = json.dumps(dict_2)
    
    with open("new_json/dict_1.json", "w") as outfile:
        outfile.write(obj_1)
     
    with open("new_json/dict_2.json", "w") as outfile:
        outfile.write(obj_2)

    # modify the rels files
    rels_list = get_rels(dict_1)
    update_rels(rels_list)
    print("AAA: ", rels_list)



if __name__ == '__main__':
    
    base_path = os.path.dirname(os.path.realpath(__file__))
    print("CURRENT_DIR:", base_path)
    sldIds = []
    dict_1 = OrderedDict()
    dict_2 = OrderedDict()
    
    sample_msg = [41, {'d': 'Onboarding','s':  [2,4,6]}, {'d': 'Presentation1','s':  [1]}]
    # sample_msg = [41, {'d': 'Onboarding','s':  [2, 4, 6]}]
    # sample_msg = [41, {'d': 'Presentation1','s':  [1]}]
    # sample_msg = [41, {'d': 'BI Case Studies','s':  [2, 3]}]

    render_id = sample_msg.pop(0)
    
    output_path = f'{base_path}/output/{str(render_id)}'
    tmp_path = f'{base_path}/tmp/{render_id}'
    input_decks = f'{base_path}/presentations'
    
    # print("TMP_PATH:", tmp_path, '\nOUT_PATH: ', output_path)  

    try:
        os.makedirs(output_path)
        os.makedirs(tmp_path)
    except:
        print("DIR ALREADY EXIST")
    
    # iterating all the messages
    deck = 1
    while sample_msg:
        deck_handler(render_id, sample_msg.pop(0), deck, dict_2)
        deck += 1

    # zipdir(f'{output_file_loc}', "Test")
