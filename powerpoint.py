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
    returns root and tree object of the xml file
    """
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(path, parser)
    root = tree.getroot()    
    return root, tree


def assetfn_to_relfn(ft, fqfn):
    """
    {ft}../customXml/item1.xml -> {ft}/customXml/item1.xml
    {ft}/ppt/slides/slide2.xml  ->  {ft}/ppt/slides/_rels/slide2.xml.rels
    """
    # if '/' in fqfn:
    l = fqfn.split("/")
    r = f"{ft}/ppt/{l[-2]}/_rels/{l[-1]}.xml.rels"
    
    return r
    


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


def build_mandatory_assets(ft, assets):
    pxr = f'{ft}/ppt/_rels/presentation.xml.rels'
    root,_ = build_tree (pxr)
    d = filter(is_manadatory, root)
    # d = filter(is_manadatory(), xml_to_dict (pxr))
    # assets = [i["@Target"] for i in d]
    
    for rel in d:
        target = rel.attrib['Target']
        if '../' in target:
            tar = f'{ft}/{target[3:]}'
        else:
            tar = f'{ft}/ppt/{target}'
        
        assets[tar] = [build_tree (tar)]
    
    return assets


def build_assets(ft, pxr, assets, slide=None):
    """
    assets = {fq-destination, file}
    req:
     assets (with all the files)
    """
    if slide:
        filename = f"{ft}/ppt/slides"
        s = f'slide{str(slide)}.xml'
        assets[f'{filename}/{s}'] = [build_tree (f'{filename}/{s}')]
        build_assets (ft, f'{filename}/_rels/{s}.rels', assets)
    
    else:
        root, tree = build_tree (pxr)
        for relation in root:
            attrib = relation.attrib
            if 'http' not in attrib['Target']:
                if '../' in attrib['Target']:
                    tar = f'{ft}/{attrib[3:]}'
                    assets[tar] = [build_tree (tar)]

                    if 'xml' in attrib['Target']:
                        fld = attrib['Target'].split('/')[-2]
                        fl = attrib['Target'].split('/')[-1]
                        pxr = f'{ft}/ppt/{fld}/_rels/{fl}'
                
                        if os.path.exists (pxr):
                            build_assets (ft, pxr, assets)
                
                else:
                    new_tar = pxr.split('/')[-3]
                    tar = f"{ft}/ppt/{new_tar}/{attrib['Target']}"
                    assets[tar] = [build_tree (tar)]
    
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
            rels[f] = [build_tree (f)]
    
    return rels


def build_content_types(ft, assets, rels, content_types):
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
        if os.path.isfile(f'{inp_path}/{i}'):
            properties[f'{inp_path}/{i}'] = [build_tree (f'{inp_path}/{i}')]
        
    return properties


def build_properties1(ft, properties):
    
    inp_path = '/'.join([ft, 'ppt'])
    out_path = f'{ft}/ppt'

    config_fls = [i for i in os.listdir(inp_path) if os.path.isfile(f'{inp_path}/{i}')]
    
    mergables = ['commentAuthors.xml', 'tableStyles.xml']
    sing_prop = ['viewProps.xml', 'presProps.xml']
    ignore = ['revisionInfo.xml']
    
    for i in config_fls:
        inp_fl = f'{inp_path}/{i}'
        out_fl = f'{out_path}/{i}'
        
        if os.path.isfile(f'{out_path}/{i}'):
            root1,tree1 = build_tree (inp_fl)
            root2,tree2 = build_tree (out_fl)
            if i in mergables:
                try:
                    for relation in [f"{root1[0].tag}"]:
                        for elt in root1.findall (relation):
                            root2.append(elt)
                except:
                    pass
            elif i in sing_prop:
                if i == 'presProps.xml':
                    inp_d = {}
                    out_lis = []
                    nm = root1.nsmap['p']
                    tag0 = f"{{{nm}}}extLst"
                    for relation in [f"{root1[0].tag}"]:
                        fp = root1.find (tag0)
                        for ele in fp:
                            attrib = ele.attrib
                            if attrib['uri'] not in inp_d.keys():
                                inp_d[attrib['uri']] = ele

                    for relation in [f"{root2[0].tag}"]:
                        fp = root2.find (tag0)
                        for ele in fp:
                            attrib = ele.attrib
                            out_lis.append (attrib['uri'])
                    for k,v in inp_d.items():
                        if k not in out_lis:
                            tag1 = root2.find(tag0)
                            tag1.append(v)
            else:
                pass
            tree2.write(out_fl, pretty_print=True, xml_declaration=True, encoding='UTF-8', standalone=True)
        else:
            shutil.copyfile(inp_fl, out_fl)


    return


def apply_assets(ft, assets):
    return assets


def apply_rels(ft, rels):
    
    return rels


def apply_content_types(ft, content_types):
    """
    need to perform refctoring of names before making changes in the output deck's file
    """
    con = '[Content_Types].xml'
    f = '/'.join ([ft, con])
    
    return content_types


def apply_properties(ft, properties):
    """
    considering 'ft' as full system putput deck's path
    """
    
    mergables = ['commentAuthors.xml', 'tableStyles.xml']
    sing_prop = ['viewProps.xml', 'presProps.xml']
    ignore = ['revisionInfo.xml']
    
    for k,v in properties.items():
        i = k.split('/')[-1]
        f = f"{ft}/ppt/{i}"
        
        if os.path.isfile (f):
            root1,_ = v[0], v[1]
            root2, tree2 = build_tree (f)
            
            if i in mergables:
                try:
                    for relation in [f"{root1[0].tag}"]:
                        for ele in root1.findall (relation):
                            root2.append(ele)
                except:
                    pass
            elif i in sing_prop:
                if i == 'presProps.xml':
                    uris = {}
                    out_lis = []
                    nm = root1.nsmap['p']
                    ext_tag = f"{{{nm}}}extLst"
                    
                    for relation in [f"{root1[0].tag}"]:
                        
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
            v[1].write(f, pretty_print=True, xml_declaration=True, encoding='UTF-8', standalone=True)
    
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

    assets = OrderedDict()
    rels = OrderedDict()
    content_types = OrderedDict()
    properties = []

    assets = build_mandatory_assets(ft, assets)

    for s in ss:
        assets = build_assets(s, ft, assets)
        rels = build_rels(ft, assets, rels)
    
    content_types = build_content_types(ft, assets, rels, content_types)
    properties = build_properties(ft, assets, properties)

    ctx["assets"] = {**ctx["assets"], assets}
    ctx["rels"] = {**ctx["rels"], rels}
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
        ctx = reduce(process_message, msgs, {})

        output_deck = f"/mnt/output/{render_id}.pptx"
        logger.debug(f"writing output deck to {output_deck}")
        write_output(ctx, output_deck)

    return ctx

# continue
