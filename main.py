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
    img = Image.open("image_adress")