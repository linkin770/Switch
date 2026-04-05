# -*- coding: utf-8 -*-
"""
配置管理模块

用于管理不同环境的配置参数，包括文件存储路径、转换参数、线程池配置等。
"""

import os
import platform

# 基础配置类
class Config:
    """基础配置类，定义所有环境共用的配置参数"""
    
    # 应用基本配置
    APP_NAME = "办公文件跨格式本地化批量转换工具"  # 应用名称
    DEBUG = True  # 调试模式
    
    # 文件存储路径
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 项目根目录
    UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")  # 上传文件存储目录
    OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")  # 输出文件存储目录
    
    # 确保目录存在
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # 创建上传目录（如果不存在）
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)  # 创建输出目录（如果不存在）
    
    # 线程池配置
    # 至少4个线程，最多CPU核心数的2倍，适用于IO密集型任务
    MAX_WORKERS = max(4, os.cpu_count() * 2)
    
    # 转换参数
    PDF_DPI = 300  # PDF转换为图像时的DPI值，影响图像清晰度
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 最大文件大小限制（100MB）
    
    # LibreOffice路径配置
    @staticmethod
    def get_libreoffice_cmd():
        """获取LibreOffice命令路径
        
        根据不同操作系统自动检测LibreOffice的安装路径
        
        Returns:
            str: LibreOffice命令的完整路径
        """
        system = platform.system()
        if system == "Windows":
            # 尝试多个可能的 LibreOffice 安装路径
            possible_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",  # 64位系统默认路径
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",  # 32位系统默认路径
                r"C:\LibreOffice\program\soffice.exe"  # 自定义安装路径
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return path
            # 如果都找不到，返回默认路径
            return r"C:\Program Files\LibreOffice\program\soffice.exe"
        else:
            # Linux系统直接使用命令名
            return "libreoffice"
    
    # 初始化LibreOffice命令路径
    LIBREOFFICE_CMD = get_libreoffice_cmd.__func__()

# 开发环境配置类
class DevelopmentConfig(Config):
    """开发环境配置，继承自基础配置类"""
    DEBUG = True  # 开发环境启用调试模式

# 生产环境配置类
class ProductionConfig(Config):
    """生产环境配置，继承自基础配置类"""
    DEBUG = False  # 生产环境禁用调试模式
    
    # 生产环境可能需要更大的线程池，以处理更多并发请求
    MAX_WORKERS = max(8, os.cpu_count() * 2)

# 根据环境变量选择配置
if os.environ.get("FLASK_ENV") == "production":
    # 如果环境变量FLASK_ENV设置为production，则使用生产环境配置
    current_config = ProductionConfig()
else:
    # 默认使用开发环境配置
    current_config = DevelopmentConfig()

# 导出配置对象，供其他模块使用
CONFIG = current_config