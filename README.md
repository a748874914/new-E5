# new-E5
from docx import Document

# 打开Word文档
doc = Document('input.docx')

# 遍历文档中的所有图片
for image in doc.inline_shapes:
    # 获取原始尺寸
    width, height = image.width, image.height

    # 计算目标缩放比例
    ratio = min(3 / width, 3 / height)

    # 缩放图片
    image.width = int(width * ratio)
    image.height = int(height * ratio)

# 保存结果
doc.save('output.docx')

