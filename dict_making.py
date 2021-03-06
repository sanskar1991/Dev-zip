import os
import shutil
import pathlib
import base
import re
import json
import natsort

from collections import OrderedDict
from lxml import etree


def list_files(fld_name, fl_name):
    """
    returns the list of files
    """
    d1 = OrderedDict()
    d2 = OrderedDict()

    lst, cnt = 0, 0
    path = f'{output_file_loc}/ppt/{fld_name}'
    
    for i in os.walk(path):
        i[2].sort(key=lambda fname: int(fname.strip(fl_name).split('.')[0]))
        lst, cnt = rename(i[2], i[0], fld_name, fl_name)
        d1.update(lst)
        d2.update(cnt)
    return d1, d2


def rename(fl_lst, path, par_fld, fl_name ):
    """
    rename a file
    """
    d_1 = OrderedDict()
    d_2 = OrderedDict()
    
    count = 1
    
    for i in fl_lst:
        fld_name = path.split('ppt/')[1]
        ext = ''.join(pathlib.Path(i).suffixes)
        # if i != f'{fl_name}{count}{ext}':
        if i != f'{fl_name}{count}':
            new_name = f'{fl_name}{count}{ext}'
            os.rename(f'{path}/{i}', f'{path}/{new_name}')
            d_1[f'{fld_name}/{i}'] = f'{fld_name}/{new_name}'
        count += 1
    d_2[par_fld] = count
    return d_1, d_2


def gen_tree(path):
    """
    generate root and tree for an xml file
    """
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    root = tree.getroot()
    return root, tree


# def modify_files(path, root, tree, name, asset_dict):
#     """
#     modify files for new assests
#     """
#     ord_dict = OrderedDict()
#     for relation in root:
#         attrib = relation.attrib
#         if name in attrib.get('Target'):
#             att = attrib.get('Target')
#             if att.split(name)[1] in asset_dict.keys():
#                 relation.set('Target', f'{name}{asset_dict[att.split(name)[1]]}')
#                 ord_dict[asset_dict.get(att.split(name)[1])] = attrib.get('Id')
                
#     tree.write(path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    
#     return ord_dict


# def update_rels(sld_dict, name):
#     """
#     update slides in rels file
#     """
#     path = f'{output_file_loc}/ppt/_rels/presentation.xml.rels'
#     root, tree = gen_tree(path)
    
#     sld = modify_files(path, root, tree, name, sld_dict)
#     return sld


# def update_med(md_dict, name, files):
#     """
#     update media in rel files of slide
#     """
#     for i in files.values():
#         path = f'{output_file_loc}/ppt/slides/_rels/{i}'
#         root, tree = gen_tree(path)
#         md = modify_files(path, root, tree, name, md_dict)
#     return md


def list_rels():
    """
    returns list of all relationship files
    """
    rel_list = []
    for i in os.walk(f'{output_file_loc}/ppt'):
        fld_name = i[0].split('ppt')[1]
        if fld_name and '_rels' in fld_name:
            for file in i[2]:
                rel_list += [f'{fld_name}/{file}']
    return rel_list


def change(file, fld_lst):
    """
    change the content of the file
    """
    path = f'{output_file_loc}/ppt/{file}'
    # print("PPPPP: ", path)
    root, tree = gen_tree(path)
    for relation in root:
        attrib = relation.attrib
        if '../' in attrib.get('Target'):
            fld = attrib.get('Target').split('../')[1]
            if fld in fld_lst:
                relation.set('Target', f'../{dict_1[fld]}')
        else:
            fld = attrib.get('Target')
            if fld in fld_lst:
                relation.set('Target', f'{dict_1[fld]}')
    tree.write(path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    return


def contents(data):
    """
    content refactoring
    """
    # list all the keys from json file
    fld_lst = data.keys()

    rel_files = list_rels()
    
    for file in rel_files:
        change(file, fld_lst)
    return
    

def max_rId():
    """
    returns maximum rId
    """
    path = f'{output_file_loc}/ppt/_rels/presentation.xml.rels'
    root, tree = gen_tree(path)
    
    rIds = []
    
    for relation in root:
        attrib = relation.attrib
        rId = int(attrib.get('Id').split('Id')[-1])
        rIds.append(rId)
        # print("ATTRIB: ", attrib)
    print("RIDSS: ", rIds)
    return max(rIds)


def copy_element(file, file_name, count, dict_3): # file = 'slides/slide1.xml'
    """
    gather all the attributes and namespace of the element
    """
    
    inp_path = f'{tmp_path}/{file_name}/'
    
    ext = ''.join(pathlib.Path(file).suffixes)
    if '../' in file:
        _,fld_name,fl_name = file.split('/')
        res = re.findall(r'(\w+?)(\d+)', fl_name)[0][0]
        new_name = f'{res}{count}{ext}'
        if os.path.exists(f'{tmp_path}/{file_name}/ppt/{file[3:]}'):
            # _,fld_name,fl_name = file.split('/')
            print("RES3 : ", re.findall(r'(\w+?)(\d+)', file[3:]))
            print("RES4 : ", re.findall(r'(\w+?)(\d+)', file[3:])[0])
            # res = re.findall(r'(\w+?)(\d+)', file[3:])[0][0]
            print("RES5 : ", res)
            print("RES6 : ", re.findall(r'(\w+?)(\d+)', file))
            shutil.copy(f'{tmp_path}/{file_name}/ppt/{file[3:]}', f"{output_file_loc}/ppt/{fld_name}/{new_name}")
            dict_3[file[3:]] = new_name
            # count += 1
            
        elif os.path.exists(f'{tmp_path}/{file_name}/{file[3:]}'):
            # new_name = f'{res}{count}{ext}'
            shutil.copy(f'{tmp_path}/{file_name}/{file[3:]}', f'{output_file_loc}/{new_name}')
            dict_3[file[3:]] = new_name
            # count += 1
    else:
        fld_name,fl_name = file.split('/')
        print("RES1 : ", re.findall(r'(\w+?)(\d+)', file))
        print("RES2 : ", re.findall(r'(\w+?)(\d+)', file)[0])
        res = re.findall(r'(\w+?)(\d+)', fl_name)[0][0]
        print("RES0 : ", res)
        new_name = f'{fld_name}/{res}{count}{ext}'
        shutil.copy(f'{tmp_path}/{file_name}/ppt/{file}', f"{output_file_loc}/ppt/{new_name}")
        dict_3[file] = new_name
        
    # count += 1
    
    return


def new_assets(target_files, file_name):
    """
    refactoring asset names
    """
    inp_path = f'{tmp_path}/{file_name}'
    
    target_files = natsort.natsorted(target_files)
    # target_files.sort(key=lambda fname: int(fname.strip(fl_name).split('.')[0]))
    dict_3 = OrderedDict()
    for i in target_files:
        if '/' in i:
            if 'slideMasters' not in i or 'slideLayouts' not in i or 'theme' not in i:
                fld_name = i.split('/')[-2]
                if fld_name in dict_2.keys():
                    count = dict_2[fld_name]
                    rId = dict_2['rId']
                else:
                    count = 1
                    rId = dict_2['rId']+1
                
                copy_element(i, file_name, count, dict_3)
                count += 1
                rId += 1
                dict_2[fld_name] = count
                dict_2['rId'] = rId
    return dict_3, dict_2
            
            
            

if __name__ == '__main__':
    base_path = os.path.dirname(os.path.realpath(__file__))
    print("CURRENT_DIR:", base_path)
    render_id = 41
    output_path = f'{base_path}/output'
    tmp_path = f'{base_path}/tmp/{render_id}'
    input_decks = f'{base_path}/presentations'
    output_file_loc = f'{output_path}/{render_id}'


    json_data = OrderedDict()
    dict_1 = OrderedDict()
    dict_2 = OrderedDict()
    m_rId = max_rId()
    dict_2.update({'rId': m_rId })
    num = -1
    
    for i in os.walk(f'{output_file_loc}/ppt/'):
        fld_name = i[0].split('ppt/')[1]
        # print("II: ", i, "\nTT: ", fld_name)
        if not fld_name:
            folders = i[1]
        else:
            if fld_name in folders:
                try:
                    res = re.findall(r'(\w+?)(\d+)', i[2][0])[0][0]
                    d1, d2 = list_files(fld_name, res)
                    dict_1.update(d1)
                    dict_2.update(d2)
                except:
                    pass
            
    obj_1 = json.dumps(dict_1)
    obj_2 = json.dumps(dict_2)
    
    with open("json/dict_1.json", "w") as outfile:
        outfile.write(obj_1)
     
    with open("json/dict_2.json", "w") as outfile:
        outfile.write(obj_2)

    contents(dict_1)
    
    target_files = ['../customXml/item1.xml', '../customXml/item2.xml', '../customXml/item3.xml', 'slideMasters/slideMaster1.xml', 'slideMasters/slideMaster2.xml', 'slideMasters/slideMaster3.xml', 'slideMasters/slideMaster4.xml', 'notesMasters/notesMaster1.xml', 'handoutMasters/handoutMaster1.xml', 'commentAuthors.xml', 'presProps.xml', 'viewProps.xml', 'theme/theme1.xml', 'tableStyles.xml', 'changesInfos/changesInfo1.xml', 'revisionInfo.xml', 'slides/slide2.xml', '../slideLayouts/slideLayout11.xml', '../slideMasters/slideMaster1.xml', '../slideLayouts/slideLayout1.xml', '../slideLayouts/slideLayout2.xml', '../media/image3.emf', '../slideLayouts/slideLayout3.xml', '../slideLayouts/slideLayout4.xml', '../slideLayouts/slideLayout5.xml', '../slideLayouts/slideLayout6.xml', '../slideLayouts/slideLayout7.xml', '../media/image4.jpeg', '../media/image5.png', '../slideLayouts/slideLayout8.xml', '../slideLayouts/slideLayout9.xml', '../slideLayouts/slideLayout10.xml', '../theme/theme1.xml', '../media/image1.emf', '../media/image2.png', '../media/image8.png', 'slides/slide4.xml', 'slides/slide6.xml', '../notesSlides/notesSlide2.xml', '../notesMasters/notesMaster1.xml', '../theme/theme5.xml', '../slides/slide6.xml', '../media/image23.png', '../media/image24.png', '../media/image25.png']

    file_name = 'Onboarding'
    a, b = new_assets(target_files, file_name)
    dict_3 = a
    obj_3 = json.dumps(dict_3)
    
    with open("json/dict_3.json", "w") as outfile:
        outfile.write(obj_3)
    
    # with open('json/dict_1.json') as f:
    #     data = json.load(f)
    # print("SSS: ", data)
    # contents(data)
        
    
    
    """
    sld = {s2:s1, s4:s2, s6:s3}
    sld[s8] = s4
    sld = {s2:s1, s4:s2, s6:s3, s8:s4}
    sld[s2] = s5
    sld = {s2:s4, s4:s2, s6:s3, s8:s4} 
    delete old dict
    count = 3
    Xml files------:
    for i in move:
        count += 1
        sld = {s2:s4}
        copy(path_of_s2, 'slides/s4')
    
    Rels files-------:
    for i in move:
        count += 1
        sld = {s2:s4}
        copy(path_of_s2, 'slides/s4')

    media_count = 8
    for i in move:
        count += 1
        sld = {s2:s4}
        copy(path_of_s2, 'slides/s4')

    master slides-----:
    slidesLayouts -------:
    theme ------:


    ***noteSlide me bhi slides use ho raha hai***



    1. rId deal
    at present : [s1, s2, s3]
    count = len(at present) = 3
    move : [s2, s3, s5]

    result : [s1, new_s2, new_s3, s5]
    during copy--
    move : [s2, s3, s5]
    count = 3
    for i in move:
        count += 1
        copy(path_of_s2, 'slides/slide{count}.xml')
    copy(source, destination)
    copy(path_of_s2, slides/)




    """
    """
    count = 0
    "slide": {}, "count": 4
    """
    """
        count = 1
        ['slide2'] = 'slide1'
        ['slide2.xml.rels'] = 'slide1.xml.rels'
        at the end count == 3
        count(json file)
    """
    """
    name changing during saving
    
    """
    
    
    
    
    
    
    
    
    