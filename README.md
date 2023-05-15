# new-E5
Sub LeftAlignColumnThree()
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
    ' 获取选中表格
    Dim tbl As Table
    Set tbl = Selection.Tables(1)
    
    ' 获取第三列，并将其设置为左对齐
    Dim col As Column
    Set col = tbl.Columns(3)
    col.SetWidth ColumnWidth:=col.Width, RulerStyle:=wdAdjustNone
    col.SetWidth ColumnWidth:=col.Width, RulerStyle:=wdAdjustProportional, _
        Param:=0.5, RelativeTo:=wdColumnWidthPoints
        ' 完成提示
    MsgBox "第三列已成功设置为左对齐。"
End Sub

