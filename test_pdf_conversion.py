# 测试PDF转换性能
import os
import time
import fitz
from pdf2docx import Converter

# 测试文件路径
pdf_path = "【津仕教育】25国考央选特招申论预测+范文.pdf"
output_path = "test_output.docx"

print(f"测试文件: {pdf_path}")
print(f"文件大小: {os.path.getsize(pdf_path) / (1024 * 1024):.2f} MB")

# 测试1: 检查PDF是否为扫描版
t1 = time.time()
doc = fitz.open(pdf_path)
num_pages = len(doc)
print(f"PDF页数: {num_pages}")

total_words = 0
for page_num in range(min(5, num_pages)):
    page = doc.load_page(page_num)
    text = page.get_text()
    total_words += len(text.split())
doc.close()
is_scanned = total_words == 0
print(f"前5页总字数: {total_words}")
print(f"是否为扫描版: {is_scanned}")
print(f"检查时间: {time.time() - t1:.2f} 秒")

# 测试2: 使用快速转换模式
t2 = time.time()
cv = Converter(pdf_path)
cv.convert(output_path, 
          start=0, 
          end=None, 
          pages=None, 
          keep_layout=False, 
          zoom=1.0, 
          table_floating=False, 
          use_ocr=False, 
          parse_structure=False)
cv.close()
print(f"快速转换时间: {time.time() - t2:.2f} 秒")

# 清理测试文件
if os.path.exists(output_path):
    os.remove(output_path)
print("测试完成")
