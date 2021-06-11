import zipfile
import os
import shutil
import xmltodict
import json


# unzip a deck
def unzip(file_path, unzip_path):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)


# zip to create a deck
def zipdir(path):
    length = len(path)
    zipf = zipfile.ZipFile('Testing.pptx', 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(path):
        folder = root[length:] # path without "parent"
        for file in files:
            zipf.write(os.path.join(root, file), os.path.join(folder, file))
    zipf.close()


# filter the directory
def ig_d(dir, files):
    return [f for f in files if f=='ppt']


# filter the files
def ig_f(dir, files):
    return [f for f in files if os.path.isfile(os.path.join(dir, f))]


# copy all the necessary files with folder architecture
def deck_handle(id, msg):
    print("DECK_HANDLE CALING...")
    file_name, slides = msg['d'], msg['s']
    unzip(dir_path+'/'+file_name+'.pptx', tmp_path+'/'+file_name)
    
    if not os.path.isdir(output_path+'/'+str(render_id)):
        # copy all the parent directories by applying directory filter
        shutil.copytree(tmp_path+'/'+file_name, output_path+'/'+str(render_id), ignore=ig_d)
        # copy all the folder architecture by applying files filter
        shutil.copytree(tmp_path+'/'+file_name+'/ppt', output_path+'/'+str(render_id)+'/ppt', ignore=ig_f)

    path = tmp_path+'/'+file_name+'/ppt/_rels/presentation.xml.rels'
    # add files in output deck
    add_files(path, file_name, slides)    
    print("FINAL: ", target)


# convert xml to dict
def xml_to_dict(path):
    print("XML_TO_DICT CALLING...")
    with open(path) as xml_file:
        data_dict = xmltodict.parse(xml_file.read())
        xml_file.close()
    if isinstance(data_dict["Relationships"]["Relationship"], list):
        data = sorted(data_dict["Relationships"]["Relationship"], key=lambda item: int(item['@Id'].split('Id')[1]))
    else:
        data = [data_dict["Relationships"]["Relationship"]]
    return data


# add files according to the relation
def add_files(path, file_name, slides=[]):
    print("ADD_FILES CALLING...")
    global target
    data = xml_to_dict(path)
    print("DATA_XML: ", data)
    # Error might be here
    if slides:
        print("IF SLIDES")
        first_slide = "slide1.xml"
        lis = []
        first_slide_id = [i["@Id"] for i in data if first_slide in i['@Target']][0].split('Id')[1]
        files = [i['@Target'] for i in data if int(first_slide_id) > int(i['@Id'].split('Id')[1])]
        print("FILES: ", files)
        print("TARGET_BEFORE: ", target)
        target = target + files
        
        print("TARGET_AFTER: ", target)
        for id in slides:
            slide = "slide"+str(id)+'.xml'
            target.append([i["@Target"] for i in data if slide in i["@Target"] and "http" not in i["@Target"]][0])
            print("AFTER_AFTER_TARGET: ", target)
            add_files(tmp_path+'/'+file_name+"/ppt/slides/_rels/"+slide+".rels",file_name)
    else:
        print("ELSE SLIDES")
        for i in data:
            if i["@Target"] in target:
                pass
            elif "http" not in i["@Target"]:
                target.append(i['@Target'])
                if ".." in i['@Target'] and "xml" in i['@Target']:
                    path = tmp_path+'/'+file_name+"/ppt/"+i['@Target'].split('..')[1].split('/')[1]+"/_rels/"+i['@Target'].split('..')[1].split('/')[2]+".rels"
                    if os.path.exists(path):
                        add_files(path, file_name)
                        
    # add the copy command
    # for i in 
    # shutil.copytree(tmp_path+'/'+file_name, output_path+'/'+str(render_id), ignore=ig_d)
                    

# recuring the files
# def file_recurse(path):
#     if "slideMaster" in path:
#         return path 
#     data = xml_to_dict(path)


if __name__ == '__main__':
    
    dir_path = os.path.dirname(os.path.realpath(__file__))
    print("CURRENT_DIR:", dir_path)
    target = []
    
    # file = open('sample_input.json')
    # sample_msg = json.load(file)
    # sample_msg = [41,{'d': 'Onboarding_1','s':  [1]}]
    sample_msg = [41,{'d': 'Presentation1','s':  [1]}]
    # file.close()

    render_id = sample_msg[0]
    output_path = "{}/output".format(dir_path)
    tmp_path = "{}/tmp/{}".format(dir_path, render_id)    
    try:
        os.makedirs(tmp_path)
        os.makedirs(output_path)
    except:
        print("DIR ALREADY EXIST")
    
    print("kkkkkkk:", tmp_path, '\nNEWWWW: ', tmp_path+'/Onboarding_1')
    for i in range(1,len(sample_msg)):
        deck_handle(render_id, sample_msg[i])

    # zipdir( tmp_path+'/Onboarding_1')
    zipdir( tmp_path+'41/Presentation1')
    # shutil.rmtree('./tmp')
