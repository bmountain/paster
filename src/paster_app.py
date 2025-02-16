from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU as p2e, pixels_to_points
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
import math
from pathlib import Path
from .data_model import Params, Children, Child

COL_WIDTH_PIXEL = 67.2  # マジックナンバー


class ExcelPaster:

    def __init__(self, params: Params, children: Children):
        self.row = 1
        self.col = 1
        self.dir_idx = 0
        self.wb = Workbook()
        self.ws = self.wb.active

        self.children: Children = children
        self.params: Params = params

    def run(self) -> None:
        """貼り付け実行"""
        self.write_all_dir()
        self.save()
        self.wb.close()

    def save(self) -> None:
        """Excelブックを保存する"""
        out = self.params.out
        sheet_name = out.replace(".xlsx", "")
        self.wb["Sheet"].title = sheet_name
        self.wb.save(out)
        print("保存しました >", out)

    def write_all_dir(self) -> None:
        """全ディレクトリの画像を張り付ける"""
        for child in self.children:
            self.write_one_dir(child)

    def write_one_dir(self, child: Child) -> None:
        """一つのディレクトリのデータを貼り付ける"""
        name = child.path.name
        img_paths = child.imgs

        self.input_cell(name)
        self.row += self.params.config.row_gap_small + 1

        height_per_row_list = []  # 一行に含まれる画像が占める行数のうち最大のもの

        for i, img in enumerate(img_paths):
            img = self.resize(img)
            self.paste_image(img)
            height_per_row, width_per_col = self.get_img_size_in_cell(img)
            height_per_row_list.append(height_per_row)

            n_cols = self.params.config.n_cols
            if i == len(img_paths) - 1:  # その項番の最後の画像の場合
                self.row += max(height_per_row_list) + self.params.config.row_gap_big
                self.col = 1
            elif i % n_cols == n_cols - 1:  # 改行するがその項番の最後の画像ではない場合
                self.row += max(height_per_row_list) + self.params.config.row_gap_small
                self.col = 1
                height_per_row_list = []
            else:  # 改行しない場合
                self.col += width_per_col + self.params.config.col_gap
        self.col = 1

    def paste_image(self, img: Image) -> None:
        """画像をWSに貼り付ける。左上を何行何列にするか指定"""
        h, w = img.height, img.width
        size = XDRPositiveSize2D(p2e(w), p2e(h))
        marker = AnchorMarker(col=self.col - 1, colOff=0, row=self.row - 1, rowOff=0)
        img.anchor = OneCellAnchor(_from=marker, ext=size)
        self.ws.add_image(img)

    def resize(self, img: Path) -> Image:
        """画像のパスを渡すとリサイズして返す"""
        img = Image(img)

        height = min(
            self.params.config.max_height,
            img.height * self.params.config.max_width / img.width,
        )
        img.height, img.width = height, img.width * height / img.height
        return img

    def get_img_size_in_cell(self, img: Image) -> tuple[int]:
        """画像を渡すとエクセルで占める行数、列数を返す（切り上げ）"""
        height_per_row = math.ceil(
            pixels_to_points(img.height, dpi=self.params.config.dpi)
            / self.params.config.row_height_point
        )
        width_per_col = math.ceil(img.width / COL_WIDTH_PIXEL)
        return height_per_row, width_per_col

    def input_cell(self, value: str) -> None:
        """WSに文字列を書き込む"""
        prefix = self.params.config.prefix
        suffix = self.params.config.suffix
        prefix = prefix if prefix is not None else ""
        suffix = suffix if suffix is not None else ""
        self.ws.cell(row=self.row, column=self.col).value = prefix + value + suffix
