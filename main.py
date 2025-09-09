"""
this program converts an excel cells to an image according to inputed image
author: amin
"""
import re
from openpyxl import Workbook
from PIL import Image

def image_on_excel(image_adress: str, excel_adress: str="", square_length: int=10):
    """it changes the cell colors acording to squares average color"""
    if not excel_adress:
        excel_adress = re.sub(r"[/\\].+", f"/sample.xslx", image_adress)
    wb = Workbook()
    ws = wb.active
    img = Image.open(image_adress)
    w, h = img.size
    for col in range(w // square_length):
        for row in range(h// square_length):
            sr = 0
            sg = 0
            sb = 0
            for i in range((col - 1)*10, col * 10):
                for j in range((row - 1)*10, row * 10):
                    r, g, b = img.getpixel((i, j))
                    sr += r
                    sg += g
                    sb += b
            t = sr / square_length, sg / square_length, sb / square_length
            print(t, (col, row))
if __name__ == "__main__":
    image_on_excel("sample.jpg", square_length=200)
