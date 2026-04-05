# -*- coding: utf-8 -*-
import os
import platform
import subprocess
from config import CONFIG

# 尝试导入python-office相关库
poexcel_available = False
poppt_available = False
poword_available = False
pandas_available = False
python_docx_available = False
weasyprint_available = False
try:
    import poexcel
    poexcel_available = True
except ImportError:
    pass
try:
    import poppt
    poppt_available = True
except ImportError:
    pass
try:
    import poword
    poword_available = True
except ImportError:
    pass
try:
    import pandas as pd
    pandas_available = True
except ImportError:
    pass
try:
    from docx import Document
    python_docx_available = True
except ImportError:
    pass
try:
    import weasyprint
    weasyprint_available = True
except ImportError:
    pass

# 使用config.py中的LibreOffice路径
LIBREOFFICE_CMD = CONFIG.LIBREOFFICE_CMD

def convert_office_to_pdf(input_path, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    
    # 检查文件扩展名
    ext = os.path.splitext(input_path)[1].lower()
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_folder, f"{base_name}.pdf")
    
    # 优先使用python-docx结合weasyprint
    if ext == ".docx" and python_docx_available and weasyprint_available:
        try:
            print("使用python-docx结合weasyprint转换DOCX为PDF")
            
            # 使用python-docx读取DOCX文件
            doc = Document(input_path)
            
            # 构建HTML内容
            html_content = f'''
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                        margin: 20px;
                        line-height: 1.5;
                    }}
                    h1, h2, h3, h4, h5, h6 {{
                        margin-top: 20px;
                        margin-bottom: 10px;
                    }}
                    p {{
                        margin-bottom: 10px;
                    }}
                    table {{
                        border-collapse: collapse;
                        width: 100%;
                        margin-bottom: 10px;
                    }}
                    th, td {{
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }}
                    th {{
                        background-color: #f2f2f2;
                    }}
                    ul, ol {{
                        margin-bottom: 10px;
                    }}
                </style>
            </head>
            <body>
            '''
            
            # 添加文档内容
            for paragraph in doc.paragraphs:
                text = paragraph.text
                if text:
                    # 检查段落是否有样式
                    style_name = paragraph.style.name
                    if style_name.startswith('Heading'):
                        # 根据标题级别生成HTML标题标签
                        level = style_name.split(' ')[1] if len(style_name.split(' ')) > 1 else '1'
                        html_content += f'<h{level}>{text}</h{level}>'
                    else:
                        html_content += f'<p>{text}</p>'
                else:
                    html_content += '<p>&nbsp;</p>'
            
            # 处理表格
            for table in doc.tables:
                html_content += '<table>'
                # 处理表头
                for row in table.rows:
                    html_content += '<tr>'
                    for cell in row.cells:
                        html_content += f'<td>{cell.text}</td>'
                    html_content += '</tr>'
                html_content += '</table>'
            
            html_content += '''
            </body>
            </html>
            '''
            
            # 使用weasyprint将HTML转换为PDF
            weasyprint.HTML(string=html_content).write_pdf(pdf_path)
            
            print(f"使用python-docx结合weasyprint成功转换DOCX为PDF: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"python-docx结合weasyprint转换失败: {e}")
    
    # 尝试使用python-office相关库
    if ext == ".docx" and poword_available:
        try:
            # 使用poword将DOCX转换为PDF
            poword.docx2pdf(path=input_path, output_path=os.path.dirname(pdf_path))
            
            # 检查生成的PDF文件是否存在
            if os.path.exists(pdf_path):
                print(f"使用poword成功转换DOCX为PDF: {pdf_path}")
                return pdf_path
            else:
                # 如果生成的PDF文件路径与期望路径不同，尝试找到它
                generated_pdf = os.path.join(os.path.dirname(pdf_path), f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf")
                if os.path.exists(generated_pdf):
                    print(f"使用poword成功转换DOCX为PDF: {generated_pdf}")
                    return generated_pdf
        except Exception as e:
            print(f"poword转换失败: {e}")
    
    elif ext in [".xls", ".xlsx"] and poexcel_available:
        try:
            # 使用poexcel将Excel转换为PDF
            print("使用poexcel转换Excel为PDF")
            # 尝试使用poexcel的excel2pdf函数
            try:
                # 尝试不同的参数组合
                try:
                    # 尝试使用第一个参数作为输入路径，第二个参数作为输出路径
                    poexcel.excel2pdf(input_path, os.path.dirname(pdf_path))
                except TypeError:
                    # 尝试使用关键字参数
                    try:
                        poexcel.excel2pdf(input_path=input_path, output_path=os.path.dirname(pdf_path))
                    except TypeError:
                        # 尝试使用其他可能的参数名
                        poexcel.excel2pdf(file=input_path, output=os.path.dirname(pdf_path))
            except AttributeError as e:
                print(f"poexcel没有excel2pdf函数: {e}")
                # 如果poexcel没有excel2pdf函数，使用comtypes作为备选方案
                try:
                    import comtypes.client
                    print("使用comtypes转换Excel为PDF")
                    
                    # 创建Excel应用程序对象
                    xlApp = comtypes.client.CreateObject("Excel.Application")
                    xlApp.Visible = False
                    xlApp.DisplayAlerts = 0
                    
                    # 打开工作簿
                    books = xlApp.Workbooks.Open(os.path.abspath(input_path), False)
                    
                    # 导出为PDF
                    books.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                    books.Close(False)
                    xlApp.Quit()
                    print("使用comtypes成功转换Excel为PDF")
                except ImportError:
                    print("comtypes不可用，尝试其他方法")
                except Exception as e:
                    print(f"comtypes转换失败: {e}")
            
            # 检查生成的PDF文件是否存在
            if os.path.exists(pdf_path):
                print(f"使用poexcel成功转换Excel为PDF: {pdf_path}")
                return pdf_path
            else:
                # 如果生成的PDF文件路径与期望路径不同，尝试找到它
                generated_pdf = os.path.join(os.path.dirname(pdf_path), f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf")
                if os.path.exists(generated_pdf):
                    print(f"使用poexcel成功转换Excel为PDF: {generated_pdf}")
                    return generated_pdf
                else:
                    print(f"PDF文件不存在: {pdf_path}")
        except ImportError:
            print("comtypes不可用，尝试其他方法")
        except Exception as e:
            print(f"poexcel转换失败: {e}")
    
    elif ext in [".ppt", ".pptx"] and poppt_available:
        try:
            # 使用poppt将PPT转换为PDF
            print("使用poppt转换PPT为PDF")
            poppt.ppt2pdf(path=input_path, output_path=os.path.dirname(pdf_path))
            
            # 检查生成的PDF文件是否存在
            if os.path.exists(pdf_path):
                print(f"使用poppt成功转换PPT为PDF: {pdf_path}")
                return pdf_path
            else:
                # 如果生成的PDF文件路径与期望路径不同，尝试找到它
                generated_pdf = os.path.join(os.path.dirname(pdf_path), f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf")
                if os.path.exists(generated_pdf):
                    print(f"使用poppt成功转换PPT为PDF: {generated_pdf}")
                    return generated_pdf
        except Exception as e:
            print(f"poppt转换失败: {e}")
    
    # 尝试使用LibreOffice
    try:
        # 确保路径是绝对路径，避免LibreOffice路径问题
        input_path = os.path.abspath(input_path)
        output_folder = os.path.abspath(output_folder)
        
        # 构建转换命令
        cmd = [LIBREOFFICE_CMD, "--headless", "--convert-to", "pdf", "--outdir", output_folder, input_path]
        
        # 执行命令并捕获输出
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"LibreOffice转换成功: {input_path}")
            return pdf_path
        else:
            print(f"LibreOffice转换失败: {result.stderr}")
    except Exception as e:
        print(f"LibreOffice转换失败: {e}")
    
    # 尝试使用Word（在Windows上）
    if platform.system() == "Windows" and ext in [".doc", ".docx"]:
        try:
            import win32com.client
            import pythoncom
            
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 创建Word应用程序对象
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            
            # 打开文档
            doc = word.Documents.Open(os.path.abspath(input_path))
            
            # 保存为PDF
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # 17 is PDF format
            doc.Close()
            word.Quit()
            
            # 释放COM资源
            pythoncom.CoUninitialize()
            
            print(f"使用Word成功转换DOCX为PDF: {pdf_path}")
            return pdf_path
        except ImportError:
            print("pywin32不可用，尝试其他方法")
        except Exception as e:
            print(f"Word转换失败: {e}")
    

    
    return None

def csv_to_xlsx(csv_path, xlsx_path):
    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
    try:
        if pandas_available:
            pd.read_csv(csv_path).to_excel(xlsx_path, index=False)
            return True
        else:
            print("pandas不可用，无法转换CSV到Excel")
            return False
    except Exception as e:
        print(f"CSV → Excel失败: {e}")
        return False

def xlsx_to_csv(xlsx_path, csv_path):
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    try:
        if pandas_available:
            # 读取Excel文件
            df = pd.read_excel(xlsx_path)
            
            # 检查是否有Unnamed列
            has_unnamed = any('Unnamed' in col for col in df.columns)
            
            # 如果有Unnamed列，为其生成默认列名
            if has_unnamed:
                # 生成默认列名，如A, B, C, ..., Z, AA, AB, ...
                def generate_column_name(index):
                    name = ''
                    while index >= 0:
                        name = chr(65 + index % 26) + name
                        index = index // 26 - 1
                    return name
                
                # 为所有列生成新的列名
                new_columns = [generate_column_name(i) for i in range(len(df.columns))]
                df.columns = new_columns
            
            # 保存为CSV文件
            df.to_csv(csv_path, index=False)
            return True
        else:
            print("pandas不可用，无法转换Excel到CSV")
            return False
    except Exception as e:
        print(f"Excel → CSV失败: {e}")
        return False