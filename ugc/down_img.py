import openpyxl
from openpyxl.drawing.image import Image
import os
import requests

file_path = 'meme.xlsx'
output_dir = f'down_img'

if not os.path.exists(output_dir):
        os.makedirs(output_dir)
wb = openpyxl.load_workbook(file_path)

sheet = wb.active

col = sheet['A']
image_index = 1
for cell in col:
    url = str(cell.value)
    if url.startswith('https'):
        response = requests.get(url)
        output_path = os.path.join(output_dir, f'{image_index}.jpg')
        if response.status_code == 200:
            with open(output_path, 'wb') as file:
                    file.write(response.content)
        image_index += 1  