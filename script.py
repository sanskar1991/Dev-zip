import zipfile
import os
import shutil
import xmltodict
import json

from pptx import Presentation
from lxml import etree, objectify


# unzip a deck
def unzip(file_path, unzip_path):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)


# zip extracted deck to get output deck
def zipdir(path):
    length = len(path)
    zipf = zipfile.ZipFile('Testing1.pptx', 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(path):
        folder = root[length:] # path without "parent"
        for file in files:
            zipf.write(os.path.join(root, file), os.path.join(folder, file))
    zipf.close()


# filter ppt folder
def ig_d(dir, files):
    return [f for f in files if f=='ppt']


# filter all files
def ig_f(dir, files):
    return [f for f in files if os.path.isfile(os.path.join(dir, f))]
    

# copy all rel files
def copy_rel(src, des):
    for x in os.walk(src):
        folder = x[0].split('ppt')[1]
        if folder and '_rels' in folder and 'slides' not in folder:
            print("SRC: ", src+folder, "\nDES: ", des+folder)
            if os.path.exists(des+folder):
                shutil.rmtree(des+folder)
            shutil.copytree(src+folder, des+folder)
    
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


# add files from input deck to output deck
def add_files(path, file_name, slides=[]):
    global target
    data = xml_to_dict(path)

    if slides:
        # get total slides
        prs = Presentation(dir_path+'/presentations/'+file_name+'.pptx')
        global tot_slides, first_slide_id
        tot_slides = len(prs.slides._sldIdLst)
        # get rId of first slide
        first_slide = "slide1.xml"
        first_slide_id = int([i["@Id"] for i in data if first_slide in i['@Target']][0].split('Id')[1])
        
        files = []
        for i in data:
            current_rId = int(i['@Id'].split('Id')[1])
            if (first_slide_id > current_rId) or (current_rId > (first_slide_id+tot_slides-1)):
                files.append(i['@Target'])
        
        print("FILESSSS: ", files)
        
        target = target + files
        for id in slides:
            slide = "slide"+str(id)+'.xml'
            target.append([i["@Target"] for i in data if slide in i["@Target"] and "http" not in i["@Target"]][0])
            shutil.copy(tmp_path+'/'+file_name+"/ppt/slides/_rels/"+slide+".rels", output_path+'/'+str(render_id)+'/ppt/slides/_rels/')
            add_files(tmp_path+'/'+file_name+"/ppt/slides/_rels/"+slide+".rels",file_name)
    else:
        for i in data:
            if i["@Target"] in target:
                pass
            elif "http" not in i["@Target"]:
                target.append(i['@Target'])
                if ".." in i['@Target'] and "xml" in i['@Target']:
                    path = tmp_path+'/'+file_name+"/ppt/"+i['@Target'].split('..')[1].split('/')[1]+"/_rels/"+i['@Target'].split('..')[1].split('/')[2]+".rels"
                    if os.path.exists(path):
                        add_files(path, file_name)
    print('TARGET: ', target)
    
    
    # copy files from tmp loc to output loc
    for i in target:
        if '../' in i:
            shutil.copy(tmp_path+'/'+file_name+'/ppt/'+i[3:], output_path+'/'+str(render_id)+'/ppt/'+i[3:].split('/')[0])
        else:
            shutil.copy(tmp_path+'/'+file_name+'/ppt/'+i, output_path+'/'+str(render_id)+'/ppt/'+i.split('/')[0])

# copy main relation file for presentation
def copy_xml_rels(path):
    # variable assigning
    global tot_slides, first_slide_id, slides
    print("SSOO: ", slides)
    slides_id = ['rId'+str(first_slide_id+i-1) for i in slides]
    # Passing the path of the xml document to enable the parsing process
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    root = tree.getroot()
    
    for relation in root:
        attrib = relation.attrib
        # print("REL_ATTR: ", attrib, type(attrib))
        print("IDDDDDD: ", attrib.get('Id'))
    
        if int(attrib.get('Id').split('Id')[1]) >= first_slide_id and int(attrib.get('Id').split('Id')[1])<(first_slide_id+tot_slides):
            if attrib.get('Id') not in slides_id:
                root.remove(relation)
                # output_path+'/'+str(render_id)+'/ppt
    tree.write(output_path+'/'+str(render_id)+'/ppt/_rels/presentation.xml.rels', pretty_print=True, xml_declaration=True, encoding='UTF-8')


# handle the deck and select files for output deck
def deck_handle(id, msg):
    global slides
    file_name, slides = msg['d'], msg['s']
    # unzip the input file
    unzip(dir_path+'/presentations/'+file_name+'.pptx', tmp_path+'/'+file_name)
    
    if not os.path.isdir(output_path+'/'+str(render_id)):
        # copy all the necessary files with folder architecture
        shutil.copytree(tmp_path+'/'+file_name, output_path+'/'+str(render_id), ignore=ig_d)
        shutil.copytree(tmp_path+'/'+file_name+'/ppt', output_path+'/'+str(render_id)+'/ppt', ignore=ig_f)
        copy_rel(tmp_path+'/'+file_name+'/ppt', output_path+'/'+str(render_id)+'/ppt')
                       
    path = tmp_path+'/'+file_name+'/ppt/_rels/presentation.xml.rels'
    add_files(path, file_name, slides)
    # global tot_slides    
    copy_xml_rels(path)
    
    print("TARGET_FILES: ", target)
    
    
if __name__ == '__main__':
    
    dir_path = os.path.dirname(os.path.realpath(__file__))
    print("CURRENT_DIR:", dir_path)
    target = []
    
    # load the message
    # file = open('sample_input.json')
    # sample_msg = json.load(file)
    # file.close()
    # sample_msg = [41,{'d': 'Onboarding_1','s':  [2,4,6]}]
    sample_msg = [41,{'d': 'Presentation1','s':  [1]}]

    render_id = sample_msg[0]
    output_path = "{}/output".format(dir_path)
    tmp_path = "{}/tmp/{}".format(dir_path, render_id)
    print("TMP_PATH:", tmp_path, '\nOUT_PATH: ', output_path)   

    try:
        os.makedirs(tmp_path)
        os.makedirs(output_path)
    except:
        print("DIR ALREADY EXIST")
    
    # iterating all the messages
    for i in range(1,len(sample_msg)):
        deck_handle(render_id, sample_msg[i])

    # zip the output deck folders
    # zipdir( tmp_path+'/Onboarding_1')
    # zipdir( tmp_path+'/Presentation1')
    # zipdir( output_path+'/Onboarding_1')
    zipdir( output_path+'/41/')
    
    # remove tmp
    # shutil.rmtree('./tmp')
