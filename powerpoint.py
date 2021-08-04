import os
import logging
import xmltodict
import shutil
import zipfile

from typing import OrderedDict
from pptx import Presentation
from functools import reduce, filter, map
from lxml import etree
# from zipfile import ZipFile
# from shutil import copyfile, rmtree, move, make_archive


logger = logging.getLogger (__name__)

fq_empty = "mob/beagle/test_resources/Empty.pptx"


def base_name (fqfn) :
    """
    returns file name, i.e. fqfn
    """
    return reduce(lambda x,f : f(x),
                  [fqfn , 
                   os.path.basename, 
                   os.path.splitext]) [0]


def file_to_work_dir (fqfn) :
    """
    returns /tmp/{fqfn}
    """
    l = ["/", "tmp", base_name (fqfn)]
    return os.path.join (*l) 


def token_to_work_dir (ft) :
    """
    """
    l = ["/", "tmp", ft]
    return os.path.join (*l) 

###

def token_is_open (ft) :
    """
    """
    fp = token_to_work_dir (ft)
    if not path.exists (fp) :
        raise ValueError (f"{ft} is not open")
    return fp


def file_is_open (fn) :
    """
    """
    fp = file_to_work_dir (fn)
    if not path.exists (fp) :
        raise ValueError (f"{ft} is not open")
    return fp


def file_is_not_open (fn) :
    """
    """
    fp = file_to_work_dir (fn)
    if os.path.exists (fp) :
        raise ValueError (f"{fn} is already open")
    return fp

###

def unpack (fqfn) :
    """
    fqfn: input deck name
    """
    fp = file_is_not_open (fqfn) # /tmp/{fqfn}
    with zipfile.ZipFile(fqfn) as z:
        z.extractall(fp)
    return fp


def pack (ft, *args, **kwargs) :
    """
    """
    fp = token_is_open (ft)
    d = kwargs.get ("file", ".".join ([ft, "pptx"]))
    z = shutil.make_archive (bn , "zip", file_to_work_dir (fqfn))
    shutil.move (z, d)
    logger.debug (f"wrote {d}")
    return d


def close (ft) :
    """
    """
    fp = token_is_open (ft)
    rmtree (fp)
    logger.debug (f"deleted {ft}")
    return


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
    unpack (d, d.split('.')[0])
    os.remove(d)
    m_rId = max_rId()
    return m_rId


def xml_to_dict (path):
    """
    convert xml file to dict
    """
    with open(path) as xml_file:
        data_dict = xmltodict.parse(xml_file.read())
        xml_file.close()
    if isinstance(data_dict["Relationships"]["Relationship"], list):
        data = sorted(data_dict["Relationships"]["Relationship"], key=lambda item: int(item['@Id'].split('Id')[1]))
    else:
        data = [data_dict["Relationships"]["Relationship"]]
    return data


def build_tree (path):
    """
    pass the path of the xml file to enable the parsing process
    """
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    root = tree.getroot()    
    return root, tree


def assetfn_to_relfn(fqfn):
    """
    {ft}/ppt/slides/slide2.xml  ->  {ft}/ppt/slides/_rels/slide2.xml.rels
    """
    l = fqfn.split("/")
    r = f"{ft}/ppt/{l[-2]}/_rels/{l[-1]}.xml"
    return 

# /mnt/input
# /mnt/output
exclude = ["slideLayouts" "slideMasters" "theme"]

def is_manadatory(i):
    """
    return true if asset should be copied
    """
    return not i["@Target"].startswith("slides") and \
        not i["@Target"].startswith("slideLayouts") and \
             not i["@Target"].startswith("slideMasters") and \
                 not i["@Target"].startswith("theme")


def build_mandatory_assets(ft, assets):
    pxr = f'{ft}/ppt/_rels/presentation.xml.rels'
    d = filter(is_manadatory(), xml_to_dict (pxr))
    assets = [i["@Target"] for i in d]

    return assets


def build_assets(slide, ft, assets):
    """
    assets = {fq-destination, file}
    req:
     assets (with all the files)

    """
    filename = f"{ft}/ppt/slides"
    return assets


def build_rels(ft, assets, rels):
    """
    add a check if the asset is having a rel file
    if yes then add it
    
    considering 'ft' as full path, "/tmp/{input_deck}
    """
    asset = ''
    for i in assets:
        if '/' in i:
            if '../' in i:
                asset = i[3:].split('/')[0]
            else:
                asset = 'ppt/' + i.split('/')[0]
            
            f = f"{i.split('/')[-1]}.rels"
            fld = f"{ft}/{asset}/_rels" 
            # fld = /tmp/{input_deck}/ppt/slides/_rels/slideX.xml.rels
            
            if os.path.isfile(f"{fld}/{f}"):
                rels.append(f)
    
    return rels


def build_content_types(ft, assets, content_types):
    """
    add content_type of new assets in [Contant_Types].xml file
    
    considering 'ft' as full system input deck path 
    """

    f = '[Content_Types].xml'
    inp_path = '/'.join ([ft, f])

    root1,_ = build_tree (inp_path)

    for relation in root1:
        if 'Override' in relation.tag:
            attrib = relation.attrib['PartName'][1:] # /ppt/slides/slide1.xml
            try:
                cnt = attrib.split ('ppt/')[-1]
                ini = '/ppt/'
            except:
                cnt = attrib
                ini = '/'
            
            if cnt in assets:
                content_types[f'{ini}{cnt}'] = relation
        else:
            attrib = relation.attrib['Extension']
            content_types[attrib] = relation
    
    return content_types


def build_properties(ft, assets, properties):
    return properties


def apply_assets(ft, assets):
    return assets


def apply_rels(ft, rels):
    return rels


def apply_content_types(ft, content_types):
    return content_types


def apply_properties(ft, properties):
    return properties


def process_message(msg, ctx):
    input_deck = msg["d"]
    
    ft = unpack (input_deck)
    
    # ft: extracted input_deck dir (/tmp/{input_deck})
    
    
    ss = msg.get("s", None)
    if not ss:
        prs = Presentation(ft)
        ss = len(prs.slides._sldIdLst)

    assets = OrderedDict()
    rels = OrderedDict()
    content_types = OrderedDict()
    properties = []

    assets = build_mandatory_deck_assets(ft, assets)

    for s in ss:
        assets = build_assets(s, ft, assets)
        rels = build_rels(ft, assets, rels)
        content_types = build_content_types(ft, assets, content_types)
        properties = build_properties(ft, assets, properties)

    ctx["assets"] = {**ctx["assets"], assets}
    ctx["rels"] = {**ctx["rels"], rels)
    ctx["content_types"] = {ctx["content_types"], rels}
    ctx["properties"] = ctx.get("properties") + properties

    return ctx

def write_output(ctx, output_deck):
    ft = new(output_deck)
    apply_assets(ft, ctx["assets"])
    apply_rels(ft, ctx["rels"])
    apply_content_types(ft, ctx["content_types"])
    apply_properties(ft, ctx["properties"])
    pack(ft, file=output_deck)
    close(ft)
    return output_deck

def render_deck_effect(msgs, fqfn):
    """
    process msgs and write output deck to fqfn
    """
    ctx = reduce(process_message, msgs, {})   
    write_output(ctx, fqfn)
    return fqfn

# continue
