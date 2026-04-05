import os
import sys

# 添加项目根目录到 Python 路径
sys.path.append('办公文件跨格式本地化批量转换工具开发V7')

from tools.office_tools import convert_office_to_pdf

# 测试 Excel 转换
excel_path = '办公文件跨格式本地化批量转换工具开发V7/uploads/FTTR 皮线安装条件评估表.xlsx'
output_folder = '办公文件跨格式本地化批量转换工具开发V7/outputs'

print(f"测试 Excel 转换: {excel_path}")
result = convert_office_to_pdf(excel_path, output_folder)
print(f"转换结果: {result}")

if result and os.path.exists(result):
    print(f"转换成功，生成的 PDF 文件: {result}")
    print(f"文件大小: {os.path.getsize(result)} 字节")
else:
    print("转换失败")
