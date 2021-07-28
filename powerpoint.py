from typing import OrderedDict
import xmltodict
from pptx import Presentation
from zipfile import ZipFile
from functools import reduce
from shutil import copyfile, rmtree, move, make_archive
import os
import logging


logger = logging.getLogger (__name__)

fq_empty = "mob/beagle/test_resources/Empty.pptx"


def base_name (fqfn) :
    """
    """
    return reduce(lambda x,f : f(x),
                  [fqfn , 
                   os.path.basename, 
                   os.path.splitext]) [0]


def file_to_work_dir (fqfn) :
    """
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
    if path.exists (fp) :
        raise ValueError (f"{fn} is already open")
    return fp

###

def unpack (fqfn) :
    """
    """
    fp = file_is_not_open (fqfn)
    with ZipFile(fqfn) as z:
        z.extractall(fp)
    return fp


def pack (ft, *args, **kwargs) :
    """
    """
    fp = token_is_open (ft)
    d = kwargs.get ("file", ".".join ([ft, "pptx"]))
    z = make_archive (bn , "zip", file_to_work_dir (fqfn))
    move (z, d)
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


# /mnt/input
# /mnt/output

def build_assets(slide, ft, assets):
    return assets

def build_rels(ft, assets, rels):
    return rels

def build_content_types(ft, assets, content_types):
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
    ft = unpack(input_deck)

    ss = msg.get("s", null)
    if not ss:
        prs = Presentation(ft)
        ss = len(prs.slides._sldIdLst)

    assets = OrderedDict()
    rels = OrderedDict()
    content_types = OrderedDict()
    properties = []

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
