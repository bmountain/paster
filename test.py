from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU as p2e, cm_to_EMU as c2e, pixels_to_points
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

import math, os, glob, datetime, re
from pathlib import Path

from app import get_file_dict

DPI = 96
IMG_HEIGHT = 400
ROW_HEIGHT_POINT = 13
COL_WIDTH_PIXEL = 67.2 #マジックナンバー
N_COL = 3 #横に並べる画像の数
COL_GAP = 1 #横に並べる画像の間で何列空けるか
ROW_GAP_SMALL = 2 #同じディレクトリ内で改行するとき何行空けるか
ROW_GAP_BIG = 3 #ディレクトリが切り替わるとき何行空けるか

class ExcelPaster:
    
    def __init__(self):
        self.row = 1
        self.col = 1
        self.dir_idx = 0
        self.wb = Workbook()
        self.ws = self.wb.active
    
    def run(self, dir_tree):
        self.write_all_dir(dir_tree)
        self.save()
        self.wb.close()
    
    def write_all_dir(self, dir_tree):
        for dir_data in dir_tree:
            self.write_one_dir(dir_data)
            self.row += ROW_GAP_BIG
    
    def save(self):
        book_name = f'{datetime.datetime.now():%Y%m%d%H%M%S}.xlsx'
        self.wb.save(book_name)
    
    def write_one_dir(self, dir_data):
        """一つのディレクトリのデータを貼り付ける"""
        dir_name = dir_data.dirpath.name
        img_paths = dir_data.imgs
        
        self.input_cell(dir_name)
        self.row += ROW_GAP_SMALL

        for i, img in enumerate(img_paths):
            img = self.get_img(img)
            self.paste_image(img)
            height_per_row, width_per_col = self.get_img_size_in_cell(img)
            if (i % N_COL == N_COL - 1) or (i == len(img_paths) - 1):
                self.row += height_per_row + ROW_GAP_SMALL
                self.col = 1
            else:
                self.col += width_per_col + COL_GAP
        self.col = 1

    def strip_extension(self, s):
        """ファイル名から拡張子を消す"""
        pattern = r'(.*)\.[^.]+'
        return re.match(pattern, s).groups()[0]

    
    def paste_image(self, img, colof=0, rowof=0):
        """画像をWSに貼り付ける。左上を何行何列にするか指定
        """

        cellh = lambda x: c2e((x * 49.77)/99)
        cellw = lambda x: c2e((x * (18.65-1.71))/10)
        coloffset = cellw(colof)
        rowoffset = cellh(rowof)

        # img = self.get_img(img)
        h, w = img.height, img.width
        size = XDRPositiveSize2D(p2e(w), p2e(h))
        marker = AnchorMarker(col=self.col, colOff=coloffset, row=self.row, rowOff=rowoffset)
        img.anchor = OneCellAnchor(_from=marker, ext=size)
        self.ws.add_image(img) 


    def get_img(self, img):
        """画像のパスを渡すとリサイズして返す"""
        img = Image(img)
        img.height, img.width = IMG_HEIGHT, img.width * IMG_HEIGHT / img.height
        return img

    def get_img_size_in_cell(self,img):
        """画像を渡すとエクセルで占める行数、列数を返す（切り上げ）"""
        height_per_row = math.ceil(pixels_to_points(img.height, dpi=DPI) / ROW_HEIGHT_POINT)
        width_per_col = math.ceil(img.width / COL_WIDTH_PIXEL)
        return height_per_row, width_per_col


    def input_cell(self, value):
        """WSに文字列を書き込む"""
        self.ws.cell(row=self.row, column=self.col).value = value

        
if __name__ == '__main__':
    xlsx_files = glob.glob('*.xlsx')
    for file in xlsx_files:
        os.remove(Path.cwd() / Path(file))
    
    dir_tree = get_file_dict('test.txt')
    paster = ExcelPaster()
    paster.run(dir_tree)
