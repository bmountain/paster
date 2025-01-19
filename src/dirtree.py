import os, re, glob
from pathlib import Path
from collections import namedtuple

class ParseException(Exception):
    pass

def get_number(s):
    """文字列が含む最後の数字を返す"""
    try:
        pattern = r"(\d+)(?!.*\d)"
        return int(re.search(pattern, s).group())
    except:
        raise ParseException('子ディレクトリに数字を含まないものがある')

def get_dirtree(dirname):
    """ディレクトリ名を渡すとその配下のディレクトリとその中の画像一覧をディレクトリツリーとして返す
"""

    cwd = Path.cwd()
    parent = Path.cwd() / Path(dirname)
    dirs = [parent / Path(child) for child in os.listdir(parent)]
    dirs.sort(key=lambda p: get_number(os.path.basename(p)))
    DirData = namedtuple('DirData', ['dirpath', 'imgs'])
    res = []
    for dir in dirs:
        dirpath = parent / dir
        os.chdir(dirpath)
        imgs = glob.glob('*.png')
        imgs.sort(key=get_number)
        imgs = [dirpath / Path(img) for img in imgs]
        dir_data = DirData(dirpath, imgs)
        res.append(dir_data)
    os.chdir(cwd)
    return res