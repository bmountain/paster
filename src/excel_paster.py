from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU as p2e, pixels_to_points
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

import math, datetime

from dirtree import get_dirtree

DPI = 96 # ディスプレイのDPI
IMG_HEIGHT = 200 # 画像の高さ
ROW_HEIGHT_POINT = 13 # EXCELの各行の高さ（ポイントで指定）
COL_WIDTH_PIXEL = 67.2 #マジックナンバー
N_COL = 3 #横に並べる画像の数
COL_GAP = 0 #横に並べる画像の間で何列空けるか
ROW_GAP_SMALL = 1 #同じディレクトリ内で改行するとき何行空けるか
ROW_GAP_BIG = 2 #ディレクトリが切り替わるとき何行空けるか

class ExcelPaster:
    
    def __init__(self, dirtree, out=None):
        self.row = 1
        self.col = 1
        self.dir_idx = 0
        self.wb = Workbook()
        self.ws = self.wb.active
        self.dirtree = dirtree
        self.out = out
    
    def run(self):
        self.write_all_dir()
        self.save()
        self.wb.close()
    
    
    def save(self):
        if self.out is None:
            book_name = f'{datetime.datetime.now():%Y%m%d%H%M%S}.xlsx'
        else:
            book_name = self.out
        self.wb.save(book_name)
        print('保存しました >', book_name)
    
    def write_all_dir(self):
        for dir_data in self.dirtree:
            self.write_one_dir(dir_data)
            self.row += ROW_GAP_BIG

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

    def paste_image(self, img):
        """画像をWSに貼り付ける。左上を何行何列にするか指定
        """
        h, w = img.height, img.width
        size = XDRPositiveSize2D(p2e(w), p2e(h))
        marker = AnchorMarker(col=self.col - 1 , colOff=0, row=self.row - 1, rowOff=0)
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