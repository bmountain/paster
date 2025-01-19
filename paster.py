from pathlib import Path
from src import excel_paster
from src import dirtree
import argparse

def parser():
    parser = argparse.ArgumentParser(description='指定した親ディレクトリ内の子ディレクトリ配下の画像をエクセルに張り付ける')
    parser.add_argument('--dirname', '-d', help='親ディレクトリ名')
    parser.add_argument('--out', '-o', help='エクセルブック名')

    args = parser.parse_args()
    return args.dirname, args.out

def main():
    dirname, out = parser()
    dir_tree = dirtree.get_dirtree(dirname)
    paster = excel_paster.ExcelPaster(dir_tree, out)
    paster.run()

if __name__ == '__main__':
    main()