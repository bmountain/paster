import openpyxl
import datetime
from pathlib import Path
from app import get_file_dict

def main():
    IMG_HEIGHT = 400

    def get_resized_img(img_path, height):
        img = openpyxl.drawing.image.Image(img_path)
        width = round(img.width * height / img.height)
        img.width = width
        img.height = height
        return img

    testimg = Path.cwd() / Path('test.png')
    testimg2 = Path.cwd() / Path('test2.png')
    img = get_resized_img(testimg, IMG_HEIGHT)
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]
    ws.add_image(img)

    book_name = f'{datetime.datetime.now():%Y%m%d%H%M%S}.xlsx'
    wb.save(book_name)

if __name__ == '__main__':
    main()