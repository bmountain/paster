import sys
sys.path.append('./src')
from src.excel_paster import ExcelPaster
from src.dirtree import get_dirtree
import argparse

def parser():
    parser = argparse.ArgumentParser(description='指定した親ディレクトリ内の子ディレクトリ配下の画像をエクセルに張り付ける')
    parser.add_argument('--dirname', '-d', help='親ディレクトリ名')
    parser.add_argument('--out', '-o', help='エクセルブック名')

    args = parser.parse_args()
    return args.dirname, args.out

def main():
    dirname, out = parser()
    dirtree = get_dirtree(dirname)
    paster = ExcelPaster(dirtree, out)
    paster.run()

if __name__ == '__main__':
    main()