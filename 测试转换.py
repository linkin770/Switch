# -*- coding: utf-8 -*-
import os
import sys

# 添加项目路径到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '办公文件跨格式本地化批量转换工具开发V7'))

from tools.office_tools import convert_office_to_pdf

# 测试DOCX到PDF转换
docx_path = r"办公文件跨格式本地化批量转换工具开发V7\uploads\头像P图教程.docx"
output_folder = r"办公文件跨格式本地化批量转换工具开发V7\outputs"

print(f"测试转换: {docx_path} → PDF")
try:
    result = convert_office_to_pdf(docx_path, output_folder)
    if result:
        print(f"转换成功！生成的PDF路径: {result}")
        if os.path.exists(result):
            print(f"PDF文件已存在，大小: {os.path.getsize(result)} 字节")
        else:
            print("错误: PDF文件生成失败")
    else:
        print("转换失败")
except Exception as e:
    print(f"转换过程中发生错误: {e}")
