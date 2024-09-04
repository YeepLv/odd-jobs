import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import os
import uuid
import json

# 读取 Excel 文件
project_id = 3
img_start_num = 1
file_path = 'test.xlsx'  # 替换为你的 Excel 文件路径
output_directory = f"output_images\\{project_id}"
if not os.path.exists(output_directory):
        os.makedirs(output_directory)
wb = openpyxl.load_workbook(file_path)

# 选择工作表
sheet = wb.active
headers = [cell.value for cell in sheet[1]]
img_path_map = {}
header_name_map = {
      '场景ID': 'scene_id',
      'AI 人物图片': 'img_url',
      '剧本名': 'script_name',
      '剧情描述': 'story_desc',
      'AI 人物描述': 'ai_character_desc',
      '用户人物描述': 'user_character_desc',
      'AI开场白': 'ai_open_speech',
      '场景状态': 'scene_status',
      '操作': 'operate',
      '不推荐理由': 'reason',
      '备注': 'remark'
}
# 遍历工作表中的所有图片
for index,img in enumerate(sheet._images):
    # 获取图片数据
    img_data = img.ref
    # 打开图像
    image = PILImage.open(img_data)

    # 保存图像到文件
    # unique_id = uuid.uuid4()
    output_path = os.path.join(output_directory, f'{img_start_num}.png')
    image.save(output_path)
    img_path_map[img_start_num] = f'{img_start_num}.png'
    print(f"Image saved as {img_start_num}")
    img_start_num += 1

row_count = 1
json_data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    item = {}
    for index,value in enumerate(row):
        key_name = header_name_map[headers[index]]
        if key_name == 'img_url':
            item['img_url'] = f'/data/local-files/?d=upload/{project_id}/{img_path_map[row_count]}'
        else:
            item[key_name] = value
    json_data.append(item)
    row_count += 1

with open('result.json', 'w') as f:
     json.dump(json_data, f)