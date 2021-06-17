import xmltodict
from zipfile import ZipFile
from functools import reduce
import shutil
# from shutil import copyfile, rmtree, move, make_archive
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


def new (ft) :
    """
    """
    fq_empty = "mob/beagle/resources/Empty.pptx"
    d = ".".join ([ft, "pptx"])
    shutil.move (fq_empty, d)
    unpack (d)
    return


