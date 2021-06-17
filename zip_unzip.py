import zipfile
import os

from functools import reduce


def unzip(file_path, unzip_path):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)


def base_name (fqfn) : # output/410.pptx
    """
    """
    return reduce(lambda x,f : f(x),
                  [fqfn , 
                   os.path.basename, 
                   os.path.splitext]) [0] # 410


def file_to_work_dir (fqfn) : # output/41.pptx
    """
    """
    l = ["/", "tmp", base_name (fqfn)] # ['/', 'tmp', '410']
    print("LLL: ", l)
    # print("ZLGLGLG: ", os.path.join (*l))
    return os.path.join (*l)  # /tmp\410


def file_is_not_open (fn) : # outpyt/41.pptx
    """
    """
    fp = file_to_work_dir (fn) # /tmp\410
    print("FILE_IS_NOT_OPEN---FPFPFPFFP: ", fp)
    if os.path.exists (fp) :
        raise ValueError (f"{fn} is already open")
    return fp


def unpack (fqfn) : # output/410.pptx
    """
    unzip the deck
    """
    fp = file_is_not_open (fqfn)
    print("UNPACK---GGHGH: ", fp)
    with zipfile.ZipFile(fqfn) as zip_ref:
        zip_ref.extractall(fp)
    return fp


# zip extracted deck to get output deck
def zipdir(path, file_name):
    length = len(path)
    zipf = zipfile.ZipFile('output/'+f'Test_{file_name}.pptx', 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(path):
        folder = root[length:] # path without "parent"
        for file in files:
            zipf.write(os.path.join(root, file), os.path.join(folder, file))
    zipf.close()