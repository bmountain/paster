import os, re, glob
from pathlib import Path
from collections import namedtuple

def get_number(s):
    """文字列が含む最後の数字を返す"""
    pattern = r"(\d+)(?!.*\d)"
    return int(re.search(pattern, s).group())

def get_file_dict(textfile):
    """テキストファイル名を渡すとそこに書かれたディレクトリ一覧を読み込む
    そのディレクトリがあることを検証する
    各ディレクトリとその中の画像すべてのパスをnamedtupleのリストとして返す
    """

    with open(textfile, 'r') as file:
        cwd = Path.cwd()
        dirs = [Path(line.rstrip()) for line in file.readlines()]
    DirData = namedtuple('DirData', ['dirpath', 'imgs'])
    res = []
    for dir in dirs:
        dirpath = cwd / dir
        assert dirpath.exists()
        os.chdir(dirpath)
        imgs = glob.glob('*.png')
        imgs.sort(key=lambda s: get_number(s))
        imgs = [dirpath / Path(img) for img in imgs]
        dir_data = DirData(dirpath, imgs)
        res.append(dir_data)
    os.chdir(cwd)
    return res

if __name__ == '__main__':
    print(get_file_dict('test.txt'))