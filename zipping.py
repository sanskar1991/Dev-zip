import zipfile
import os

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

# path = 'tmp/41/presentation1'
path = 'tmp/41/Onboarding'
prep_path = 'presentations/Onboarding.pptx'
# unzip(prep_path, path)
zipdir(path, "Test1")