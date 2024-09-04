import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import os
import uuid

# 读取 Excel 文件
file_path = 'test.xlsx'  # 替换为你的 Excel 文件路径
output_directory = "output_images"
if not os.path.exists(output_directory):
        os.makedirs(output_directory)
wb = openpyxl.load_workbook(file_path)

# 选择工作表
sheet = wb.active
# 遍历工作表中的所有图片
for img in sheet._images:
    # 获取图片数据
    img_data = img.ref
    print(img.path)
    # 打开图像
    image = PILImage.open(img_data)

    # 保存图像到文件
    unique_id = uuid.uuid4()
    output_path = os.path.join(output_directory, f'{unique_id}.png')
    image.save(output_path)
    print(f"Image saved as {unique_id}")

print("所有图片已保存。")