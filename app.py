# -*- coding: utf-8 -*-
import os
import shutil
from flask import Flask, request, jsonify, send_from_directory
from concurrent.futures import ThreadPoolExecutor, as_completed
from tools import pdf_to_docx, txt_to_pdf, markdown_to_pdf, latex_to_pdf, convert_office_to_pdf, csv_to_xlsx, xlsx_to_csv, make_zip

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/convert', methods=['POST'])
def convert():
    target_format = request.form.get('format', '').lower()
    files = request.files.getlist('files')

    if not files:
        return jsonify({"status":"error","msg":"请先选择文件"})

    shutil.rmtree(UPLOAD_FOLDER, ignore_errors=True)
    shutil.rmtree(OUTPUT_FOLDER, ignore_errors=True)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # 保存所有文件
    saved_files = []
    for file in files:
        filename = file.filename
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        os.makedirs(os.path.dirname(input_path), exist_ok=True)
        
        # 如果文件已经存在，先删除它
        if os.path.exists(input_path):
            try:
                os.remove(input_path)
                print(f"已删除已存在的文件: {input_path}")
            except Exception as e:
                print(f"删除已存在的文件失败: {e}")
        
        # 保存新文件
        file.save(input_path)
        saved_files.append(filename)

    logs, success, fail = [], 0, 0

    # 定义转换函数
    def process_file(filename):
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        ext = os.path.splitext(filename)[1].lower().replace(".", "")
        relative_path = os.path.relpath(input_path, UPLOAD_FOLDER)
        out_ext = "pdf" if target_format == "pdf" else target_format

        # 同名文件自动加原格式区分：1.doc → 1_doc.pdf
        file_basename = os.path.splitext(relative_path)[0]
        new_basename = f"{file_basename}_{ext}"
        out_path = os.path.join(OUTPUT_FOLDER, f"{new_basename}.{out_ext}")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        ok = False
        log_message = ""

        pdf_conversion = {
            "doc": convert_office_to_pdf,
            "docx": convert_office_to_pdf,
            "rtf": convert_office_to_pdf,
            "html": convert_office_to_pdf,
            "xml": convert_office_to_pdf,
            "xls": convert_office_to_pdf,
            "xlsx": convert_office_to_pdf,
            "ppt": convert_office_to_pdf,
            "pptx": convert_office_to_pdf,
            "txt": txt_to_pdf,
            "md": markdown_to_pdf,
            "tex": lambda src,dst: latex_to_pdf(src, os.path.dirname(dst))
        }

        try:
            if target_format == "pdf":
                func = pdf_conversion.get(ext)
                if func:
                    if ext in ["doc", "docx", "rtf", "html", "xml", "xls", "xlsx", "ppt", "pptx"]:
                        generated_pdf = func(input_path, os.path.dirname(out_path))
                        ok = generated_pdf is not None and os.path.exists(generated_pdf)
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
                    else:
                        ok = func(input_path, out_path)
                else:
                    log_message = f"⚠️ 不支持转换: {filename} → {target_format}"
                    return log_message, False
            elif target_format in ["xlsx","xls"]:
                if ext == "csv":
                    ok = csv_to_xlsx(input_path, out_path)
                else:
                    log_message = f"⚠️ 不支持转换: {filename} → {target_format}"
                    return log_message, False
            elif target_format == "csv":
                if ext in ["xlsx","xls"]:
                    ok = xlsx_to_csv(input_path, out_path)
                else:
                    log_message = f"⚠️ 不支持转换: {filename} → {target_format}"
                    return log_message, False
            elif target_format == "docx":
                if ext == "pdf":
                    # 检查PDF是否为扫描版
                    import fitz
                    doc = fitz.open(input_path)
                    total_words = 0
                    for page_num in range(min(3, len(doc))):
                        page = doc.load_page(page_num)
                        text = page.get_text()
                        total_words += len(text.split())
                    doc.close()
                    is_scanned = total_words == 0
                    
                    ok = pdf_to_docx(input_path, out_path)
                    if not ok:
                        if is_scanned:
                            log_message = f"⚠️ 警告：扫描版PDF转换失败，请检查文件是否损坏: {filename} → {new_basename}.{out_ext}"
                        else:
                            log_message = f"❌ 失败：{filename} → {new_basename}.{out_ext}"
                        return log_message, False
                else:
                    log_message = f"⚠️ 不支持转换: {filename} → {target_format}"
                    return log_message, False
            else:
                log_message = f"⚠️ 不支持转换: {filename} → {target_format}"
                return log_message, False

            if ok and os.path.exists(out_path):
                log_message = f"✅ 成功：{filename} → {new_basename}.{out_ext}"
                return log_message, True
            else:
                log_message = f"❌ 失败：{filename} → {new_basename}.{out_ext}"
                return log_message, False
        except Exception as e:
            log_message = f"❌ 失败：{filename} → {new_basename}.{out_ext}，错误：{e}"
            return log_message, False

    # 使用线程池并行处理文件
    # 增加线程池大小，对于IO密集型任务，可以使用更多线程
    max_workers = max(4, os.cpu_count() * 2)  # 至少4个线程，最多CPU核心数的2倍
    print(f"使用线程池，最大工作线程数: {max_workers}")
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        future_to_file = {executor.submit(process_file, filename): filename for filename in saved_files}
        
        # 收集结果
        for future in as_completed(future_to_file):
            log_message, is_success = future.result()
            logs.append(log_message)
            if is_success:
                success += 1
            else:
                fail += 1

    zip_path = os.path.join(OUTPUT_FOLDER, "converted_files.zip")
    make_zip(OUTPUT_FOLDER, zip_path)

    return jsonify({
        "status": "ok",
        "success": success,
        "fail": fail,
        "logs": "\n".join(logs),
        "zip_url": "/download_zip/converted_files.zip"
    })


@app.route('/outputs/<path:filename>')
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

@app.route('/download_zip/<path:filename>')
def download_zip(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

@app.route('/clear_temp', methods=['POST'])
def clear_temp():
    shutil.rmtree(UPLOAD_FOLDER, ignore_errors=True)
    shutil.rmtree(OUTPUT_FOLDER, ignore_errors=True)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    return jsonify({"status":"ok","msg":"临时文件已清空"})

if __name__=="__main__":
    app.run(host='127.0.0.1', port=5000, debug=True)