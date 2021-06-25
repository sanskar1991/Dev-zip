import os
import shutil
import pathlib
import base
import re
import json

from collections import OrderedDict
from lxml import etree


def list_files(fld_name, fl_name):
    """
    returns the list of files
    """
    path = f'{output_file_loc}/ppt/{fld_name}'
    print("PATH: ", path)
    print("FLD_NAME: ", fld_name)
    file_list = []
    for i in os.walk(path):
        i[2].sort(key=lambda fname: int(fname.strip(fl_name).split('.')[0]))
        lst, cnt = rename(i[2], i[0], fl_name)
        file_list.append(lst)
    file_list.append(cnt)
    return file_list


def rename(fl_lst, path, fl_name):
    """
    rename a file
    """
    file_dict = OrderedDict()
    count = 1
    for i in fl_lst:
        if i != f'{fl_name}{count}':
            ext = ''.join(pathlib.Path(i).suffixes)
            file_dict[i] = f'{fl_name}{count}{ext}'
            os.rename(f'{path}/{i}', f'{path}/{fl_name}{count}{ext}')
        count += 1
    return file_dict, {"count": count}


def gen_tree(path):
    """
    generate root and tree for an xml file
    """
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    root = tree.getroot()
    return root, tree


def modify_files(path, root, tree, name, asset_dict):
    """
    modify files for new assests
    """
    ord_dict = OrderedDict()
    for relation in root:
        attrib = relation.attrib
        if name in attrib.get('Target'):
            att = attrib.get('Target')
            if att.split(name)[1] in asset_dict.keys():
                relation.set('Target', f'{name}{asset_dict[att.split(name)[1]]}')
                ord_dict[asset_dict.get(att.split(name)[1])] = attrib.get('Id')
                
    tree.write(path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    
    return ord_dict


def update_rels(sld_dict, name):
    """
    update slides in rels file
    """
    path = f'{output_file_loc}/ppt/_rels/presentation.xml.rels'
    root, tree = gen_tree(path)
    
    sld = modify_files(path, root, tree, name, sld_dict)
    return sld


def update_med(md_dict, name, files):
    """
    update media in rel files of slide
    """
    for i in files.values():
        path = f'{output_file_loc}/ppt/slides/_rels/{i}'
        root, tree = gen_tree(path)
        md = modify_files(path, root, tree, name, md_dict)
    return md
    

if __name__ == '__main__':
    base_path = os.path.dirname(os.path.realpath(__file__))
    print("CURRENT_DIR:", base_path)
    render_id = 41
    output_path = f'{base_path}/output'
    tmp_path = f'{base_path}/tmp/{render_id}'
    input_decks = f'{base_path}/presentations'
    output_file_loc = f'{output_path}/{render_id}'
        

    # media = OrderedDict()
    # slides = OrderedDict()
    # layouts = OrderedDict()
    # masters = OrderedDict()

    json_data = []
    num = -1
    
    for i in os.walk(f'{output_file_loc}/ppt/'):
        fld = OrderedDict()
        fld_name = i[0].split('ppt/')[1]
        if not fld_name:
            folders = i[1]
            print("FOLDER: ", folders)
            num += 1
        else:
            if fld_name in folders:
            # if folders[num] not in fld.keys():
                try:
                    res = re.findall(r'(\w+?)(\d+)', i[2][0])[0][0]
                    print("RES: ", res)
                    fl_lis = list_files(folders[num], res)
                    print("fl_lis: ", fl_lis)
                    fld[folders[num]] = fl_lis
                    print("HHH: ", fld)
                    json_data.append(fld)
                except:
                    pass
                
                
                num += 1
            
    obj = json.dumps(json_data)
    
    with open("sample.json", "w") as outfile:
        outfile.write(obj)
            
    
    # dir_list = {'media': 'image', 'slides': 'slide', 'slideMasters': 'slideMaster', 'slideLayouts': 'slideLayout', 'theme': 'theme'}
    
    # for i, j in dir_list.items():
    #     if 'media' == i:
    #         media = list_files(i, j)
    #         # update_med(media[0], 'media/')
    #         # print("MEDIA: ", media)
    #     elif 'slides' == i:
    #         slides = list_files(i, j)
    #         sld = update_rels(slides[0], 'slides/')
    #         md = update_med(media[0] , 'media/', slides[1])
    #         # print("MDD: ", md, "\nSLD: ", sld)
    #         print("SLIDES: ", slides, "\nHHH: ", sld)
    #     elif 'slideMasters' == i:
    #         masters = list_files(i, j)
    #         # print("MASTERS: ", masters)
    #         print()
    #     elif 'slideLayouts' == i:
    #         layouts = list_files(i, j)
    #         # print("LAYOUTS: ", layouts)
    #         print()
    #     elif 'theme' == i:
    #         theme = list_files(i, j)
    #         # print("THEME: ", theme)
    #         print()
    #         # print("THEME: ", theme)

    # a = [OrderedDict([('image1.emf', 'image1.emf'), ('image2.png', 'image2.png'), ('image3.emf', 'image3.emf'), ('image4.jpeg', 'image4.jpeg'), ('image5.png', 'image5.png'), ('image8.png', 'image6.png'), ('image23.png', 'image7.png'), ('image24.png', 'image8.png'), ('image25.png', 'image9.png')])]
    
    # # print("SLIDES: ", slide_list)
    
    # count = 1
    # # for 
    # # for i in slide_list:
    
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
    
    
    
    
    
    
    
    
    