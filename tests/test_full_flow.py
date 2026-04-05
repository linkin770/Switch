# -*- coding: utf-8 -*-
import os
import sys

# 添加项目路径到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '办公文件跨格式本地化批量转换工具开发V7'))

from tools.office_tools import convert_office_to_pdf

# 测试完整的转换流程，模拟app.py中的调用方式
docx_path = r"办公文件跨格式本地化批量转换工具开发V7\uploads\头像P图教程.docx"
output_folder = r"办公文件跨格式本地化批量转换工具开发V7\outputs"

# 构建输出路径
ext = os.path.splitext(os.path.basename(docx_path))[1].lower().replace(".", "")
file_basename = os.path.splitext(os.path.basename(docx_path))[0]
new_basename = f"{file_basename}_{ext}"
out_path = os.path.join(output_folder, f"{new_basename}.pdf")

print(f"测试完整转换流程: {docx_path} → {out_path}")
try:
    # 调用convert_office_to_pdf函数
    generated_pdf = convert_office_to_pdf(docx_path, os.path.dirname(out_path))
    
    # 检查转换是否成功
    ok = generated_pdf is not None and os.path.exists(generated_pdf)
    
    # 如果成功且生成的PDF路径与期望路径不同，重命名
    if ok and generated_pdf != out_path:
        # 如果目标文件存在，先删除它
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except Exception as e:
                print(f"删除已存在的文件失败: {e}")
        
        # 重命名文件，添加重试机制
        max_retries = 3
        retry_count = 0
        while retry_count < max_retries:
            try:
                os.rename(generated_pdf, out_path)
                print(f"重命名PDF文件: {generated_pdf} → {out_path}")
                break
            except Exception as e:
                print(f"重命名文件失败: {e}")
                retry_count += 1
                # 等待一段时间后重试
                import time
                time.sleep(1)
        
        # 如果所有重试都失败，使用原始路径
        if retry_count >= max_retries:
            print("所有重命名尝试都失败，使用原始路径")
            out_path = generated_pdf
    
    # 检查最终输出文件是否存在
    if os.path.exists(out_path):
        print(f"转换成功！生成的PDF路径: {out_path}")
        print(f"PDF文件大小: {os.path.getsize(out_path)} 字节")
    else:
        print("转换失败：最终PDF文件不存在")
except Exception as e:
    print(f"转换过程中发生错误: {e}")
