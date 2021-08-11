import os
import logging
import xmltodict
import shutil
import zipfile
import pathlib
import re

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
    if not os.path.exists (fp) :
        raise ValueError (f"{ft} is not open")
    return fp


def file_is_open (fn) :
    """
    """
    fp = file_to_work_dir (fn)
    if not os.path.exists (fp) :
        raise ValueError (f"{fp} is not open")
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

    fqfn : Onboarding
    fp = /tmp/Onboarding 

    fqfn : f"{base_path}/{presentation}/Onboarding.pptx"

    /mnt/input/{messsage-fn}
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
    shutil.rmtree (fp)
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
    
    # m_rId = max_rId()
    
    return "some output path"


def max_rId (emp):
    """
    find the largest rId
    """
    path = f'{emp}/ppt/_rels/presentation.xml.rels'
    root, tree = build_tree(path)

    rIds = []

    for relation in root:
        attrib = relation.attrib
        rId = int(attrib.get('Id').split('Id')[-1])
        rIds.append(rId)
    return {'rId': max(rIds)}


def get_fld_f_names (asset):
    """
    generates folder and file names
    """
    sp = asset.split ('/')
    if '_rels' in asset: # slides/_rels/slide2.xml.rels
        f = sp[-1]
        fld = f'{sp[-3]}/{sp[-2]}'
    else:
        fld, f = sp[-2], sp[-1]

    return fld, f


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


def create(tree, f):
    """
    creates an XML file using the tree
    """
    tree.write(f, pretty_print=True, xml_declaration=True, encoding='UTF-8', standalone=True)
    
    return tree


def build_tree (path):
    """
    pass the path of the xml file to enable the parsing process
    returns root and tree object of the xml file
    """
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    # root = tree.getroot()    
    return tree


def build_name(ft, k, v, refactored_cnt):
    """
    generates refactored name for an asset,
    update the refactored count and
    creates assets in the output deck
    """
    fld, f = get_fld_f_names(k)
    assets_nm, assets_cnt = {}, {}

    ext = ''.join(pathlib.Path(f).suffixes) # .xml (or) .xml.rels
    name = re.findall(r'(\w+?)(\d+)', f)[0][0] # slide, slideLayout, theme, slideMaster

    if f'{fld}/{name}' in refactored_cnt.keys():
        cnt = refactored_cnt[f'{fld}/{name}']+1
    else:
        cnt = 1
    
    new_name = f'{name}{cnt}{ext}'
    tree = v
    
    if 'ppt' in k:
        create(tree, f"{ft}/ppt/{fld}/{new_name}")
    else:
        create(tree, f"{ft}/{fld}/{new_name}")

    assets_nm[f'{fld}/{f}'] = f'{fld}/{new_name}'
    assets_cnt[f'{fld}/{name}'] = cnt
    
    return assets_nm, assets_cnt


def generate_last_indicies(ft):
    """
    generate refactored count dict for assets
    ft: path_to_output_deck
    """
    for i in os.walk(ft):
        pass
        


def assetfn_to_relfn(ft, fqfn):
    """
    {ft}../customXml/item1.xml -> {ft}/customXml/item1.xml
    {ft}/ppt/slides/slide2.xml  ->  {ft}/ppt/slides/_rels/slide2.xml.rels
    """
    # if '/' in fqfn:
    l = fqfn.split("/")
    r = f"{ft}/ppt/{l[-2]}/_rels/{l[-1]}.xml.rels"
    
    return r


def walk_asset_tree(ft, asset, coll={}):
    
    rf = assetfn_to_relfn(ft, asset)

    assetfn = f"{ft}/{asset}"
    coll = {**coll, assetfn: build_tree(assetfn)}

    if os.path.exists(rf):
        # rels = open rels
        # for r in rels:
            # t = target
            # coll = walk_asset_tree(t, coll)
        pass

    return coll

# /mnt/input
# /mnt/output
exclude = ["slideLayouts" "slideMasters" "theme"]

def is_manadatory(i):
    """
    return true if asset should be copied
    """
    attrib = i.attrib['Target']
    return '/' in attrib and \
        not attrib.startswith("slides") and \
            not attrib.startswith("slideLayouts") and \
                not attrib.startswith("slideMasters") and \
                    not attrib.startswith("theme")
                 
    # return not i["@Target"].startswith("slides") and \
    #     not i["@Target"].startswith("slideLayouts") and \
    #          not i["@Target"].startswith("slideMasters") and \
    #              not i["@Target"].startswith("theme")


def next_index_for(type, lasts):
    """
    
    """
    current = lasts.get(type, 0)
    current = current + 1
    lasts[type] = current
    return (current, lasts)


def build_mandatory_assets(ft, assets):
    pxr = f'{ft}/ppt/_rels/presentation.xml.rels'
    tree = build_tree (pxr)
    root = tree.getroot()
    d = filter(is_manadatory, root)
    # d = filter(is_manadatory(), xml_to_dict (pxr))
    # assets = [i["@Target"] for i in d]
    
    for rel in d:
        target = rel.attrib['Target']
        if '../' in target:
            tar = f'{ft}/{target[3:]}'
        else:
            tar = f'{ft}/ppt/{target}'
        
        assets[tar] = build_tree (tar)
    
    return assets

#def build_assets(s, ft, assets)
def build_assets(ft, pxr, assets, slide=None):
    """
    assets = {fq-destination, file}
    req:
     assets (with all the files)
    """
    # s = f'slide{str(slide)}.xml'
    # assets = walk_asset_tree(ft, s)

    if slide:
        filename = f"{ft}/ppt/slides"
        s = f'slide{str(slide)}.xml'
        f = f'{ft}/{filename}/{s}'
        assets[f] = build_tree (f)
        build_assets (ft, f'{filename}/_rels/{s}.rels', assets)
    
    else:
        tree = build_tree (pxr)
        root = tree.getroot()
        for relation in root:
            attrib = relation.attrib
            if 'http' not in attrib['Target']:
                if '../' in attrib['Target']:
                    tar = f'{ft}/{attrib[3:]}'
                    assets[tar] = build_tree (tar)

                    if 'xml' in attrib['Target']:
                        fld = attrib['Target'].split('/')[-2]
                        fl = attrib['Target'].split('/')[-1]
                        pxr = f'{ft}/ppt/{fld}/_rels/{fl}'
                
                        if os.path.exists (pxr):
                            build_assets (ft, pxr, assets)
                
                else:
                    new_tar = pxr.split('/')[-3]
                    tar = f"{ft}/ppt/{new_tar}/{attrib['Target']}"
                    assets[tar] = build_tree (tar)
    
    return assets


def build_rels(ft, assets, rels):
    """
    add a check if the asset is having a rel file
    if yes then add it
    
    considering 'ft' as full path, "/tmp/{input_deck}
    """
    for i in assets.keys():
        f = assetfn_to_relfn(ft, i)         
        if os.path.isfile(f):
            rels[f] = build_tree (f)
    
    return rels


def build_content_types(ft, assets, rels, content_types):
    """
    add content_type of new assets in [Contant_Types].xml file
    
    considering 'ft' as full system input deck path 
    """

    f = '[Content_Types].xml'
    inp_path = '/'.join ([ft, f])

    tree1 = build_tree (inp_path)
    root1 = tree1.getroot()

    for relation in root1:
        if 'Override' in relation.tag:
            attrib = relation.attrib['PartName'][1:] # /ppt/slides/slide1.xml
            try:
                cnt = attrib.split ('ppt/')[-1]
                ini = f'{ft}/ppt'
            except:
                cnt = attrib
                ini = f'{ft}'
            
            name = f'{ini}/{cnt}'
            
            if name in assets.keys() or name in rels.keys():
                content_types[name] = relation
        else:
            attrib = relation.attrib['Extension']
            content_types[attrib] = relation
    
    return content_types


def build_properties(ft, properties):
    
    inp_path = '/'.join([ft, 'ppt'])
    
    for i in os.listdir(inp_path):
        f = f'{inp_path}/{i}'
        if os.path.isfile(f):
            properties[f] = build_tree (f)
        
    return properties


def apply_assets(ft, assets, rels, refactored_cnt):
    """
    creates all assets in the output deck 
    """
    for k,v in assets.items():
        assets_nm, assets_cnt = build_name (ft, k, v, refactored_cnt)
    return assets_nm, assets_cnt


def refactor_content(ft, refactored_nm, ):
    """
    refactoring content of the rel files
    and then saving it
    """




def apply_rels(ft, k, v, refactored_nm):
    """
    creates all rel files of assets in the output deck 
    """
    fld, f = get_fld_f_names(k)
    assets_nm, assets_cnt = {}, {}

    ext = ''.join(pathlib.Path(f).suffixes) # .xml (or) .xml.rels
    name = re.findall(r'(\w+?)(\d+)', f)[0][0] # slide, slideLayout, theme, slideMaster

    if f'{fld}/{name}' in assets_cnt.keys():
        cnt = assets_cnt[f'{fld}/{name}']+1
    else:
        cnt = 1
    
    new_name = f'{name}{cnt}{ext}'
    tree = v
    
    if 'ppt' in k:
        create(tree, f"{ft}/ppt/{fld}/{new_name}")
    else:
        create(tree, f"{ft}/{fld}/{new_name}")

    assets_nm[f'{fld}/{f}'] = f'{fld}/{new_name}'
    assets_cnt[f'{fld}/{name}'] = cnt
    
    return assets_nm, assets_cnt
    


# def apply_rels(ft, rels, refactored_nm, refactored_cnt):
#     """
#     creates all the rel files of assets in the output deck
#     """
#     for k,v in rels:
#         assets_nm, assets_cnt = refactor_content (ft, k, v, refactored_nm)
#     return assets_nm, assets_cnt


def apply_content_types(ft, content_types):
    """
    need to perform refctoring of names before making changes in the output deck's file
    """
    con = '[Content_Types].xml'
    f = '/'.join ([ft, con])
    
    return content_types


def apply_properties(ft, properties):
    """
    considering 'ft' as full system output deck's path
    """
    
    mergables = ['commentAuthors.xml', 'tableStyles.xml']
    sing_prop = ['viewProps.xml', 'presProps.xml']
    ignore = ['revisionInfo.xml']
    
    for k,v in properties.items():
        i = k.split('/')[-1]
        f = f"{ft}/ppt/{i}"
        
        if os.path.isfile (f):
            tree1 = v
            root1 = tree1.getroot()
            tree2 = build_tree (f)
            root2 = tree2.getroot()
            
            etag = f"{root1[0].tag}"
            if i in mergables:
                try:
                    for relation in [etag]:
                        for ele in root1.findall (relation):
                            root2.append(ele)
                except IndexError:
                    print("list index out of range")
            elif i in sing_prop:
                if i == 'presProps.xml':
                    uris = {}
                    out_lis = []
                    nm = root1.nsmap['p']
                    ext_tag = f"{{{nm}}}extLst"
                    
                    for relation in [etag]:
                        
                        fp = root1.find (ext_tag)
                        for ele in fp:
                            attrib = ele.attrib
                            if attrib['uri'] not in uris.keys():
                                uris[attrib['uri']] = ele

                    for relation in [f"{root2[0].tag}"]:
                        fp = root2.find (ext_tag)
                        for ele in fp:
                            attrib = ele.attrib
                            out_lis.append (attrib['uri'])
                    
                    for k,v in uris.items():
                        if k not in out_lis:
                            tag1 = root2.find(ext_tag)
                            tag1.append(v)

            tree2.write(f, pretty_print=True, xml_declaration=True, encoding='UTF-8', standalone=True)
        
        else:
            tree1.write(f, pretty_print=True, xml_declaration=True, encoding='UTF-8', standalone=True)
    
    return properties


def process_message(msg, ctx):
    logger.debug(f"processing message {msg}")
    input_deck = f"/mnt/input/{msg['d']}"
    
    ft = unpack (input_deck)
    
    # ft: extracted input_deck dir (/tmp/{input_deck})
    
    
    ss = msg.get("s", None)
    if not ss:
        prs = Presentation(ft)
        ss = list(range(1, len(prs.slides._sldIdLst) + 1))

    lasts = ctx["lasts"]
    assets = OrderedDict()
    rels = OrderedDict()
    content_types = OrderedDict()
    properties = OrderedDict()
    ref_names = dict()
    ref_count = dict()

    assets = build_mandatory_assets(ft, assets)

    pxr = ''

    for s in ss:
        assets = build_assets(ft, pxr, assets, lasts, s)
        rels = build_rels(ft, assets, rels, lasts)
    
    content_types = build_content_types(ft, assets, rels, content_types)
    properties = build_properties(ft, assets, properties)

    ctx["assets"] = {**ctx["assets"], **assets}
    ctx["rels"] = {**ctx["rels"], **rels}
    ctx["content_types"] = {ctx["content_types"], rels}
    ctx["properties"] = ctx.get("properties") + properties

    ctx["lasts"] = lasts
    return ctx

def write_output(ft, ctx, output_deck):
    
    ref_nm, ref_cnt = apply_assets(ft, ctx["assets"], ctx['refactored_cnt'])
    ctx['refactored_name'] = {**ref_nm}
    ctx['refactored_count'] = {**ref_nm}
    
    ref_nm, ref_cnt = apply_rels(ft, ctx["rels"], ctx['refactored_name'], ctx['refactored_cnt'])
    ctx['refactored_name'] = {**ref_nm}
    ctx['refactored_count'] = {**ref_nm}
    
    apply_content_types(ft, ctx["content_types"])
    apply_properties(ft, ctx["properties"])
    
    pack(ft, file=output_deck)
    close(ft)
    return output_deck

def render_deck_effect(ctx):
    """
    process msgs and write output deck to fqfn

    fqfn should be /mnt/output/{render_id}
    """
    msgs = ctx.get ("msgs", [])
    if not msgs:
        logger.error("No messages in context")
    else:
        render_id = msgs.pop(0)
        output_deck = f"/mnt/output/{render_id}.pptx"
        
        # ft: full system path of empty output deck
        ft = new(output_deck)
        refactored_cnt = {**max_rId(ft)}
        
        lasts = generate_last_indicies(ft)
        
        ctx = reduce(process_message, msgs, {"lasts": lasts})

        
        logger.debug(f"writing output deck to {output_deck}")
        write_output(ft, ctx, output_deck, refactored_cnt)

    return ctx

# continue

