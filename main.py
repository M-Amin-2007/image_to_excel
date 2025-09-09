"""
this program converts an excel cells to an image according to inputed image
author: amin
"""
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from PIL import Image

def number_to_letters(n):
    result = ""
    while n > 0:
        n -= 1
        result = chr(ord('A') + n % 26) + result
        n //= 26
    return result

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
            rgb_color = sr // square_length, sg // square_length, sb // square_length
            hex_color = f"{rgb_color[0]:02X}{rgb_color[1]:02X}{rgb_color[2]:02X}"
            color = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            print(col + 1, number_to_letters(col + 1))
            cell_id = f"{number_to_letters(col + 1)}{row + 1}"
            print(cell_id)
            ws[cell_id].fill = color

if __name__ == "__main__":
    image_on_excel("sample.jpg", square_length=200)
